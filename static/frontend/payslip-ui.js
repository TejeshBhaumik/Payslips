const form = document.querySelector("#payslipForm");
const excelFile = document.querySelector("#excelFile");
const options = document.querySelector("#options");
const customFields = document.querySelector("#customFields");
const emailOnlyFields = document.querySelector("#emailOnly");
const websiteField = document.querySelector("#field16");
const submitBtn = document.querySelector("#submitBtn");
const submitLabel = document.querySelector("#submitBtn .submit-label");
const submitStatus = document.querySelector("#submitStatus");
const progressPanel = document.querySelector("#progressPanel");
const progressMessage = document.querySelector("#progressMessage");
const progressCount = document.querySelector("#progressCount");
const progressCurrent = document.querySelector("#progressCurrent");
const progressTrack = document.querySelector(".progress-track");
const progressBar = document.querySelector("#progressBar");
const runState = document.querySelector(".run-state");
const limitModal = document.querySelector("#limitModal");
const limitModalTitle = document.querySelector("#limitModalTitle");
const limitModalMessage = document.querySelector("#limitModalMessage");
const limitModalClose = document.querySelector("#limitModalClose");
const successModal = document.querySelector("#successModal");
const successModalClose = document.querySelector("#successModalClose");

let pollTimer = null;
const alertedWaitStates = new Set();

const workflowCopy = {
    default: "Choose a workflow before starting the batch.",
    custom: "Custom payslip copy enabled.",
    salarySlip: "Salary slip template selected.",
    emailOnly: "Email-only campaign fields enabled.",
};

function setFieldsDisabled(container, disabled) {
    container.querySelectorAll("input, select, textarea").forEach((input) => {
        input.disabled = disabled;
    });
}

function setFormDisabled(disabled) {
    form.querySelectorAll("input, select, textarea").forEach((input) => {
        input.disabled = disabled;
    });
}

function setSectionVisibility(section, visible) {
    section.hidden = !visible;
    setFieldsDisabled(section, !visible);
}

function setRunState(label, tone = "ready") {
    runState.lastChild.textContent = ` ${label}`;
    runState.dataset.tone = tone;
}

function updateVisibleFields() {
    const selectedOption = options.value;
    const isCustom = selectedOption === "custom";
    const isEmailOnly = selectedOption === "emailOnly";

    setSectionVisibility(customFields, isCustom);
    setSectionVisibility(emailOnlyFields, isEmailOnly);

    websiteField.required = isEmailOnly;
    websiteField.disabled = !isEmailOnly;
    submitStatus.textContent = workflowCopy[selectedOption] || workflowCopy.default;
}

function updateFileLabel(input) {
    const label = document.querySelector(`[data-file-label="${input.name}"]`);
    if (!label) {
        return;
    }

    label.textContent = input.files.length ? input.files[0].name : label.dataset.defaultText;
}

function cacheDefaultFileLabels() {
    document.querySelectorAll("[data-file-label]").forEach((label) => {
        label.dataset.defaultText = label.textContent;
    });
}

function setProgress(job) {
    const total = Number(job.total || 0);
    const sent = Number(job.sent || 0);
    const percent = total > 0 ? Math.min(100, Math.round((sent / total) * 100)) : 0;

    progressPanel.hidden = false;
    progressMessage.textContent = job.message || "Preparing batch.";
    progressCount.textContent = `${sent} / ${total}`;
    progressCurrent.textContent = job.current ? `Current: ${job.current}` : "";
    progressTrack.setAttribute("aria-valuenow", String(percent));
    progressBar.style.width = `${percent}%`;
}

function clearPollTimer() {
    clearTimeout(pollTimer);
    pollTimer = null;
}

function resetSubmitButton(label = "Start batch") {
    submitBtn.disabled = false;
    submitBtn.classList.remove("is-loading");
    submitLabel.textContent = label;
    setFormDisabled(false);
    updateVisibleFields();
}

function showLimitModal(message) {
    limitModalTitle.textContent = "batch stopped at email limit";
    limitModalMessage.textContent = message;
    limitModal.hidden = false;
    limitModalClose.focus();
}

function hideLimitModal() {
    limitModal.hidden = true;
}

function showSuccessModal() {
    successModal.hidden = false;
    successModalClose.focus();
}

function hideSuccessModal() {
    successModal.hidden = true;
}

async function readJsonResponse(response) {
    const contentType = response.headers.get("content-type") || "";
    if (contentType.includes("application/json")) {
        return response.json();
    }

    return { error: await response.text() };
}

async function pollJob(jobId) {
    const response = await fetch(`/extract/status/${jobId}`, {
        headers: { Accept: "application/json" },
    });
    const job = await readJsonResponse(response);

    if (!response.ok) {
        throw new Error(job.error || "Could not read batch status.");
    }

    setProgress(job);

    if (job.status === "complete") {
        clearPollTimer();
        resetSubmitButton("Start another batch");
        submitStatus.textContent = "all payslips generated and sent";
        progressMessage.textContent = "all payslips generated and sent";
        setRunState("Complete", "success");
        showSuccessModal();
        return { done: true };
    }

    if (job.status === "limited") {
        clearPollTimer();
        const sent = Number(job.sent || 0);
        const total = Number(job.total || 0);
        const retryText = job.message || "Email provider stopped sending. Please retry later.";
        const limitMessage = total > 0
            ? `Stopped after sending ${sent} of ${total}. ${retryText}`
            : retryText;
        resetSubmitButton("Retry batch");
        submitStatus.textContent = limitMessage;
        progressMessage.textContent = limitMessage;
        setRunState("Limit hit", "warning");
        showLimitModal(limitMessage);
        return { done: true };
    }

    if (job.status === "waiting") {
        const sent = Number(job.sent || 0);
        const total = Number(job.total || 0);
        const retryAt = job.retry_at ? new Date(job.retry_at) : null;
        const retryDelay = retryAt ? Math.max(1000, retryAt.getTime() - Date.now() + 5000) : 600000;
        const retryTime = retryAt ? retryAt.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }) : "soon";
        const waitMessage = total > 0
            ? `Stopped after sending ${sent} of ${total}. Retrying automatically at ${retryTime} from Excel row ${job.resume_row}.`
            : `Retrying automatically at ${retryTime}.`;

        submitStatus.textContent = waitMessage;
        progressMessage.textContent = waitMessage;
        setRunState("Waiting", "warning");
        showLimitModal(waitMessage);
        const waitStateKey = `${job.id}:${job.retry_at || ""}:${job.resume_row || ""}`;
        if (!alertedWaitStates.has(waitStateKey)) {
            alertedWaitStates.add(waitStateKey);
            window.alert(waitMessage);
        }
        return { done: false, delayMs: retryDelay };
    }

    if (job.status === "failed") {
        clearPollTimer();
        resetSubmitButton("Retry batch");
        submitStatus.textContent = job.error || "Batch failed.";
        setRunState("Failed", "error");
        return { done: true };
    }

    return { done: false, delayMs: 1000 };
}

function schedulePoll(jobId, delayMs) {
    pollTimer = window.setTimeout(async () => {
        try {
            const result = await pollJob(jobId);
            if (!result.done) {
                schedulePoll(jobId, result.delayMs || 1000);
            }
        } catch (error) {
            clearPollTimer();
            submitStatus.textContent = error.message;
            setRunState("Failed", "error");
            resetSubmitButton("Retry batch");
        }
    }, delayMs);
}

async function startPolling(jobId) {
    const result = await pollJob(jobId);
    if (!result.done && !pollTimer) {
        schedulePoll(jobId, result.delayMs || 1000);
    }
}

async function handleSubmit(event) {
    event.preventDefault();

    if (!excelFile.files.length) {
        window.alert("please enter employee data excel");
        submitStatus.textContent = "please enter employee data excel";
        excelFile.focus();
        return;
    }

    if (!form.checkValidity()) {
        form.reportValidity();
        submitStatus.textContent = "Complete the required fields before starting.";
        return;
    }

    const formData = new FormData(form);

    clearPollTimer();
    hideLimitModal();
    hideSuccessModal();
    setFormDisabled(true);
    submitBtn.disabled = true;
    submitBtn.classList.add("is-loading");
    submitLabel.textContent = "Processing";
    submitStatus.textContent = "Starting batch.";
    setRunState("Running", "running");
    setProgress({ sent: 0, total: 0, message: "Uploading workbook.", current: "" });

    try {
        const response = await fetch(form.action, {
            method: "POST",
            body: formData,
            headers: { Accept: "application/json" },
        });
        const payload = await readJsonResponse(response);

        if (!response.ok) {
            throw new Error(payload.error || "Batch could not be started.");
        }

        submitStatus.textContent = "Batch started. Tracking payslips as they send.";
        await startPolling(payload.jobId);
    } catch (error) {
        submitStatus.textContent = error.message;
        setRunState("Failed", "error");
        resetSubmitButton("Retry batch");
    }
}

function initPayslipUi() {
    if (!form || !options) {
        return;
    }

    cacheDefaultFileLabels();
    updateVisibleFields();

    options.addEventListener("change", updateVisibleFields);
    form.addEventListener("submit", handleSubmit);
    limitModalClose.addEventListener("click", hideLimitModal);
    successModalClose.addEventListener("click", hideSuccessModal);

    form.querySelectorAll("input[type='file']").forEach((input) => {
        input.addEventListener("change", () => updateFileLabel(input));
    });

    const jobId = new URLSearchParams(window.location.search).get("jobId");
    if (jobId) {
        setFormDisabled(true);
        submitBtn.disabled = true;
        submitBtn.classList.add("is-loading");
        submitLabel.textContent = "Processing";
        submitStatus.textContent = "Batch started. Tracking payslips as they send.";
        setRunState("Running", "running");
        setProgress({ sent: 0, total: 0, message: "Loading batch status.", current: "" });
        startPolling(jobId);
        window.history.replaceState({}, "", window.location.pathname);
    }
}

initPayslipUi();
