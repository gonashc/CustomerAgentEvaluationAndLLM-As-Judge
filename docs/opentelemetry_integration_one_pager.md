# OpenTelemetry Integration Guide

## What This Adds

The Week 4 notebook evaluates a customer support classifier. It runs tickets through a baseline prompt, compares predictions against known labels, improves the prompt, and measures whether the new version performs better.

OpenTelemetry adds trace-level visibility to that workflow.

In plain English: every ticket classification becomes a traceable event. That event records what the model was asked to do, what it predicted, what the correct answer was, and whether the prediction was right.

## Vocabulary: What Is A Span?

A `span` is one traceable unit of work.

In this project, one span is created each time the classifier handles one customer ticket. If the notebook classifies 100 tickets, OpenTelemetry can create 100 spans.

Each span acts like a timestamped record for one ticket classification. It can store the ticket id, prompt version, expected label, predicted label, correctness, model reasoning, timing, and status.

A `trace` is the larger story that can contain one or more spans. For this Week 4 use case, the easiest way to explain it is:

- span = one ticket classification
- trace = the broader record that lets you inspect what happened

## OpenTelemetry vs. LangSmith

LangSmith is the place where you inspect AI runs, prompts, model calls, and eval results.

OpenTelemetry is the standard way the notebook emits trace data.

They work together here:

- OpenTelemetry captures structured spans from the notebook.
- LangSmith receives those spans and gives you a UI for reviewing them.

So OpenTelemetry does not replace LangSmith. It makes the tracing format more portable and standard, while LangSmith remains the place where the traces are easiest to inspect.

## Where It Was Added

The main integration is in the ticket prediction loop. Each time the classifier processes a ticket, the notebook creates a span named:

```text
customer_support.classify_ticket
```

That span includes:

- ticket id
- run name, such as baseline or improved
- prompt version
- true support category
- predicted support category
- whether the prediction was correct
- model reasoning

This connects each row in the CSV results back to a trace in LangSmith.

## Why It Helps Week 4

The Week 4 assignment is about more than getting a higher score. It asks you to understand the model's mistakes, group failures, change the prompt, and compare the before-and-after results.

OpenTelemetry helps because it gives evidence for that story:

- You can inspect a failed ticket instead of only looking at aggregate accuracy.
- You can compare baseline and improved runs using the same metadata.
- You can explain why a prompt change helped one failure category.
- You can also see where the change caused regressions.

## Endpoint Note

If the notebook prints:

```text
Failed to export span batch code: 404, reason: Not Found
```

that means the classifier ran, but traces were being sent to an incorrect OpenTelemetry endpoint.

The notebook now points OpenTelemetry to LangSmith's OTLP trace endpoint:

```text
https://api.smith.langchain.com/otel/v1/traces
```

That fixes the trace export issue and keeps the eval workflow connected to LangSmith.

## Short Talk Track

"I used LangSmith as the AI eval and trace-inspection workspace, and OpenTelemetry as the standard tracing layer. Each customer ticket classification creates a span with the ticket id, prompt version, expected label, predicted label, correctness, and model reasoning. This lets me connect the accuracy results back to specific model traces, which makes the Week 4 failure analysis easier to explain."
