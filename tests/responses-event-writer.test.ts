import { describe, expect, test } from "bun:test";
import { classifyBridgeFailure } from "../src/proxy/failure-classifier";
import {
  ResponsesEventWriter,
  type ResponsesTerminalState,
} from "../src/proxy/responses-event-writer";
import type { JsonObject } from "../src/proxy/types";

describe("ResponsesEventWriter", () => {
  test("stamps a mixed event stream with contiguous sequence numbers", () => {
    const events: JsonObject[] = [];
    const writer = new ResponsesEventWriter((event) => events.push(event));
    const response = baseResponse("in_progress");

    writer.created(response);
    writer.inProgress(response);
    writer.outputItemAdded("resp_1", 0, messageItem("item_1"));
    writer.contentPartAdded("resp_1", 0, "item_1", outputPart());
    writer.outputTextDelta("resp_1", 0, "item_1", "hello");
    writer.outputTextDone("resp_1", 0, "item_1", "hello");
    writer.contentPartDone("resp_1", 0, "item_1", outputPart());
    writer.outputItemDone("resp_1", 0, messageItem("item_1"));
    writer.functionCallArgumentsDelta("resp_1", 1, "item_2", "{");
    writer.functionCallArgumentsDone("resp_1", 1, "item_2", "{}");
    writer.customToolInputDelta("resp_1", 2, "item_3", "patch");
    writer.customToolInputDone("resp_1", 2, "item_3", "patch");
    writer.webSearchCallInProgress("resp_1", 3, "search_1");
    writer.webSearchCallSearching("resp_1", 3, "search_1");
    writer.webSearchCallCompleted("resp_1", 3, "search_1");
    writer.completed(baseResponse("completed"));

    expect(events.map((event) => event.sequence_number)).toEqual(
      events.map((_, index) => index),
    );
    expect(new Set(events.map((event) => event.sequence_number)).size).toBe(
      events.length,
    );
    expect(events.at(-1)?.type).toBe("response.completed");
  });

  test("rejects every semantic emission after a terminal", () => {
    const completedEvents: JsonObject[] = [];
    const completed = new ResponsesEventWriter((event) =>
      completedEvents.push(event),
    );
    completed.completed(baseResponse("completed"));
    expect(() => completed.created(baseResponse("in_progress"))).toThrow();
    expect(() => completed.completed(baseResponse("completed"))).toThrow();
    expect(completedEvents).toHaveLength(1);

    const failed = new ResponsesEventWriter(() => {});
    failed.failed(
      baseResponse("in_progress"),
      classifyBridgeFailure("upstream_timeout"),
    );
    expect(() => failed.completed(baseResponse("completed"))).toThrow();

    const doubleCompleted = new ResponsesEventWriter(() => {});
    doubleCompleted.completed(baseResponse("completed"));
    expect(() => doubleCompleted.completed(baseResponse("completed"))).toThrow();
  });

  test("emits completed, failed, and incomplete response shapes", () => {
    const completedEvents: JsonObject[] = [];
    const completed = new ResponsesEventWriter((event) =>
      completedEvents.push(event),
    );
    completed.completed(baseResponse("in_progress"));
    expect(completedEvents[0]).toMatchObject({
      type: "response.completed",
      response: {
        status: "completed",
        error: null,
        incomplete_details: null,
      },
    });

    const failedEvents: JsonObject[] = [];
    const failed = new ResponsesEventWriter((event) => failedEvents.push(event));
    const timeout = classifyBridgeFailure("upstream_timeout");
    failed.failed(baseResponse("in_progress"), timeout);
    expect(failedEvents[0]).toMatchObject({
      type: "response.failed",
      response: {
        id: "resp_1",
        created_at: 1_700_000_000,
        model: "gpt-5.6-sol",
        usage: { total_tokens: 3 },
        output: [],
        status: "failed",
        error: {
          code: "upstream_timeout",
          message: "The provider did not respond before the deadline.",
        },
        incomplete_details: null,
      },
    });
    expect(failedEvents[0]?.response).toEqual({
      ...baseResponse("failed"),
      status: "failed",
      error: {
        code: "upstream_timeout",
        message: "The provider did not respond before the deadline.",
      },
      incomplete_details: null,
    });

    const incompleteEvents: JsonObject[] = [];
    const incomplete = new ResponsesEventWriter((event) =>
      incompleteEvents.push(event),
    );
    incomplete.incomplete(
      baseResponse("in_progress"),
      classifyBridgeFailure("partial_or_unprovable_completion"),
    );
    expect(incompleteEvents[0]).toMatchObject({
      type: "response.incomplete",
      response: {
        id: "resp_1",
        created_at: 1_700_000_000,
        model: "gpt-5.6-sol",
        usage: { total_tokens: 3 },
        output: [],
        status: "incomplete",
        error: null,
        incomplete_details: { reason: "partial_or_unprovable_completion" },
      },
    });
    expect(incompleteEvents[0]?.response).toEqual({
      ...baseResponse("incomplete"),
      status: "incomplete",
      error: null,
      incomplete_details: { reason: "partial_or_unprovable_completion" },
    });
  });

  test("records client abort without emitting a wire terminal", () => {
    const events: JsonObject[] = [];
    const writer = new ResponsesEventWriter((event) => events.push(event));

    writer.clientAbort();

    expect(events).toEqual([]);
    expect(writer.terminalState).toEqual({ kind: "cancelled" });
    expect(() =>
      writer.incomplete(
        baseResponse("incomplete"),
        classifyBridgeFailure("partial_or_unprovable_completion"),
      ),
    ).toThrow();
    expect(writer.terminalState?.kind).toBe("cancelled");
  });

  test("exposes the selected terminal state", () => {
    const writer = new ResponsesEventWriter(() => {});
    writer.completed(baseResponse("completed"));

    const state: ResponsesTerminalState | null = writer.terminalState;
    expect(state?.kind).toBe("completed");
    expect(state?.response.status).toBe("completed");
  });
});

function baseResponse(status: string): JsonObject {
  return {
    id: "resp_1",
    object: "response",
    created_at: 1_700_000_000,
    status,
    error: null,
    incomplete_details: null,
    model: "gpt-5.6-sol",
    usage: { total_tokens: 3 },
    output: [],
    output_text: "",
    parallel_tool_calls: false,
    store: false,
  };
}

function messageItem(id: string): JsonObject {
  return {
    id,
    type: "message",
    status: "in_progress",
    role: "assistant",
    content: [],
  };
}

function outputPart(): JsonObject {
  return { type: "output_text", text: "", annotations: [] };
}
