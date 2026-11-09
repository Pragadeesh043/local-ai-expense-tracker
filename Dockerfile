FROM python:3.11-slim AS builder

ENV VIRTUAL_ENV=/opt/venv
ENV PATH="${VIRTUAL_ENV}/bin:${PATH}"
RUN python -m venv "$VIRTUAL_ENV"

WORKDIR /build
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

FROM python:3.11-slim AS runtime

ENV VIRTUAL_ENV=/opt/venv \
    PATH="/opt/venv/bin:${PATH}" \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    EXPENSE_XLSX=/data/expenses.xlsx

WORKDIR /app
COPY --from=builder /opt/venv /opt/venv
COPY expense_tracker/ ./expense_tracker/

VOLUME ["/data"]
ENTRYPOINT ["python", "-m", "expense_tracker"]
CMD ["check"]
