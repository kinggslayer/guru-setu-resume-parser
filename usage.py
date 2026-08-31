"""
Track what a run costs.

Running out of credit mid-batch is a confusing failure — every resume
fails with a 429 and nothing explains why. Showing spend per run makes the
balance something you watch rather than discover.

Prices are per MILLION tokens, in USD, for the model in OPENAI_MODEL.
They are a published figure that can change, so the display is always
labelled an estimate. USD_TO_INR is likewise indicative.
"""

import threading


# gpt-4.1-mini, USD per million tokens.
PRICE_PER_MILLION = {
    "input": 0.40,
    "output": 1.60,
}

USD_TO_INR = 88.0


class UsageTracker:
    """
    Token totals for one run.

    Thread-safe because extraction runs on a worker pool — without the lock
    two threads finishing together can lose an update.
    """

    def __init__(self):
        self._lock = threading.Lock()
        self.reset()

    def reset(self):

        with self._lock:
            self.input_tokens = 0
            self.output_tokens = 0
            self.calls = 0

    def record(self, response):
        """
        Add one API response. Silently ignores responses with no usage
        block — a missing count should never break a parse.
        """

        usage = getattr(response, "usage", None)

        if usage is None:
            return

        with self._lock:
            self.input_tokens += getattr(usage, "prompt_tokens", 0) or 0
            self.output_tokens += getattr(usage, "completion_tokens", 0) or 0
            self.calls += 1

    @property
    def usd(self):

        return (
            self.input_tokens / 1_000_000 * PRICE_PER_MILLION["input"]
            + self.output_tokens / 1_000_000 * PRICE_PER_MILLION["output"]
        )

    @property
    def inr(self):
        return self.usd * USD_TO_INR

    def summary(self):
        """One line for the UI, or None when nothing was spent."""

        if not self.calls:
            return None

        return (
            f"{self.calls} API call(s) — "
            f"{self.input_tokens:,} input + {self.output_tokens:,} output "
            f"tokens, roughly ${self.usd:.3f} (about Rs {self.inr:.2f}). "
            "Estimate based on published gpt-4.1-mini pricing."
        )


# One tracker per process; reset at the start of each run.
tracker = UsageTracker()
