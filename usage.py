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
#
# "cached" applies to the repeated instruction prefix. OpenAI caches
# prompt prefixes over 1024 tokens automatically and bills them at a
# quarter of the normal input rate — which is why the instructions live
# in a separate system message, byte-identical on every call.
PRICE_PER_MILLION = {
    "input": 0.40,
    "cached": 0.10,
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
            self.cached_tokens = 0
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

        # Cached prefix tokens are still counted in prompt_tokens, but are
        # billed at a quarter of the rate. Ignoring them overstated the
        # cost of every run.
        details = getattr(usage, "prompt_tokens_details", None)
        cached = getattr(details, "cached_tokens", 0) or 0 if details else 0

        with self._lock:
            self.input_tokens += getattr(usage, "prompt_tokens", 0) or 0
            self.cached_tokens += cached
            self.output_tokens += getattr(usage, "completion_tokens", 0) or 0
            self.calls += 1

    @property
    def usd(self):

        fresh = max(self.input_tokens - self.cached_tokens, 0)

        return (
            fresh / 1_000_000 * PRICE_PER_MILLION["input"]
            + self.cached_tokens / 1_000_000 * PRICE_PER_MILLION["cached"]
            + self.output_tokens / 1_000_000 * PRICE_PER_MILLION["output"]
        )

    @property
    def inr(self):
        return self.usd * USD_TO_INR

    def summary(self):
        """One line for the UI, or None when nothing was spent."""

        if not self.calls:
            return None

        per_resume = self.inr / self.calls if self.calls else 0

        cached_note = ""

        if self.cached_tokens:
            share = 100 * self.cached_tokens / max(self.input_tokens, 1)
            cached_note = (
                f" {self.cached_tokens:,} of the input tokens ({share:.0f}%) "
                "were served from cache at a quarter of the price."
            )

        return (
            f"{self.calls} API call(s) — "
            f"{self.input_tokens:,} input + {self.output_tokens:,} output "
            f"tokens, roughly ${self.usd:.3f} (about Rs {self.inr:.2f}, "
            f"Rs {per_resume:.2f} per resume).{cached_note} "
            "Estimate based on published gpt-4.1-mini pricing."
        )


# One tracker per process; reset at the start of each run.
tracker = UsageTracker()
