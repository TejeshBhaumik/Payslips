# Agent Notes

## Frontend Progress Tracking

The `/extract` endpoint used to be synchronous:

```text
browser submits form -> Flask sends all payslips -> Flask returns one final response
```

That shape cannot show live frontend progress because the browser receives no response until all payslips are done.

The current approach uses a background job plus polling:

```text
browser POST /extract -> Flask creates jobId -> Flask starts background thread -> Flask returns jobId
browser GET /extract/status/<jobId> every second -> frontend updates progress bar
background thread sends payslips -> updates job progress after each send
```

## Why Threads Are Used

Creating a thread object does not start work:

```python
thread = threading.Thread(target=run_extract_job, args=(...))
```

Calling `start()` begins a second execution path:

```python
thread.start()
```

After `thread.start()`:

- The original Flask request thread returns `{"jobId": job_id}` quickly.
- The background thread runs `run_extract_job(...)` and keeps generating/sending payslips.
- Later polling requests read progress from `/extract/status/<job_id>`.

Calling `run_extract_job(...)` directly would not be equivalent. It would run on the current request thread and the browser would wait until the whole batch finished.

## Why The Lock Exists

Progress is stored in the in-memory `JOBS` dictionary.

The background worker writes progress:

```python
update_job(job_id, sent=3, message="Sent payslip 3 of 12")
```

The frontend polling endpoint reads progress:

```python
GET /extract/status/<job_id>
```

Because those can happen at the same time, `JOBS_LOCK` makes reads and writes take turns. This avoids inconsistent access to shared state.

## User-Facing States

The frontend should stay on the first page and show:

- A loading/progress bar while payslips are being generated and sent.
- Per-payslip progress such as `Sent payslip 3 of 12`.
- `all payslips generated and sent` when the job completes.
- A popup saying how many messages were sent when the SMTP provider stops the batch because of throttling or quota.

## Rate Limit Behavior

`sendEmail.emailDelivery()` raises `RateLimitExceeded` only when the SMTP provider returns a throttling or quota-style response.

That exception is intentionally surfaced as a job status:

```json
{
  "status": "limited",
  "message": "Email provider stopped sending. Please retry later."
}
```

The frontend detects `status === "limited"` and shows the modal.

## Production Caveat

The current `JOBS` store is in memory. That is fine for a small single-process Flask app, but it is not durable and is not shared across multiple Gunicorn worker processes.

If this app needs multi-worker production reliability, move job state to Redis, a database, or a real task queue such as Celery/RQ.
