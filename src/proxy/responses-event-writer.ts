import {
  buildResponseContentPartAddedEvent,
  buildResponseContentPartDoneEvent,
  buildResponseCompletedEvent,
  buildResponseCreatedEvent,
  buildResponseCustomToolInputDeltaEvent,
  buildResponseCustomToolInputDoneEvent,
  buildResponseFailedEvent,
  buildResponseFunctionCallArgumentsDeltaEvent,
  buildResponseFunctionCallArgumentsDoneEvent,
  buildResponseIncompleteEvent,
  buildResponseInProgressEvent,
  buildResponseOutputItemAddedEvent,
  buildResponseOutputItemDoneEvent,
  buildResponseOutputTextDeltaEvent,
  buildResponseOutputTextDoneEvent,
  buildResponseWebSearchCallCompletedEvent,
  buildResponseWebSearchCallInProgressEvent,
  buildResponseWebSearchCallSearchingEvent,
} from "./responses-api";
import {
  type BridgeFailure,
} from "./failure-classifier";
import type { JsonObject } from "./types";
import { cloneJsonValue } from "./utils";

export type ResponsesEventSink = (event: JsonObject) => void;

export type ResponsesTerminalState =
  | { kind: "completed"; response: JsonObject }
  | { kind: "failed"; response: JsonObject; failure: BridgeFailure }
  | {
      kind: "incomplete";
      response: JsonObject;
      failure: BridgeFailure;
    }
  | { kind: "cancelled" };

export class ResponsesEventWriter {
  constructor(private readonly sink: ResponsesEventSink) {}

  private nextSequence = 0;
  private terminal: ResponsesTerminalState | null = null;

  get sequenceNumber(): number {
    return this.nextSequence;
  }

  get terminalState(): ResponsesTerminalState | null {
    return this.terminal;
  }

  get isTerminal(): boolean {
    return this.terminal !== null;
  }

  emit(event: JsonObject): JsonObject {
    this.assertOpen();
    if (
      event.type === "response.completed" ||
      event.type === "response.failed" ||
      event.type === "response.incomplete"
    ) {
      throw new Error(
        "Terminal response events must be emitted through completed, failed, or incomplete",
      );
    }
    return this.emitEvent(event);
  }

  created(response: JsonObject): JsonObject {
    return this.emit(buildResponseCreatedEvent(response));
  }

  inProgress(response: JsonObject): JsonObject {
    return this.emit(buildResponseInProgressEvent(response));
  }

  outputItemAdded(
    responseId: string,
    outputIndex: number,
    item: JsonObject,
  ): JsonObject {
    return this.emit(
      buildResponseOutputItemAddedEvent(responseId, outputIndex, item),
    );
  }

  outputItemDone(
    responseId: string,
    outputIndex: number,
    item: JsonObject,
  ): JsonObject {
    return this.emit(
      buildResponseOutputItemDoneEvent(responseId, outputIndex, item),
    );
  }

  contentPartAdded(
    responseId: string,
    outputIndex: number,
    itemId: string,
    part: JsonObject,
    contentIndex = 0,
  ): JsonObject {
    return this.emit(
      buildResponseContentPartAddedEvent(
        responseId,
        outputIndex,
        itemId,
        part,
        contentIndex,
      ),
    );
  }

  contentPartDone(
    responseId: string,
    outputIndex: number,
    itemId: string,
    part: JsonObject,
    contentIndex = 0,
  ): JsonObject {
    return this.emit(
      buildResponseContentPartDoneEvent(
        responseId,
        outputIndex,
        itemId,
        part,
        contentIndex,
      ),
    );
  }

  outputTextDelta(
    responseId: string,
    outputIndex: number,
    itemId: string,
    delta: string,
    contentIndex = 0,
  ): JsonObject {
    return this.emit(
      buildResponseOutputTextDeltaEvent(
        responseId,
        outputIndex,
        itemId,
        delta,
        contentIndex,
      ),
    );
  }

  outputTextDone(
    responseId: string,
    outputIndex: number,
    itemId: string,
    text: string,
    contentIndex = 0,
  ): JsonObject {
    return this.emit(
      buildResponseOutputTextDoneEvent(
        responseId,
        outputIndex,
        itemId,
        text,
        contentIndex,
      ),
    );
  }

  functionCallArgumentsDelta(
    responseId: string,
    outputIndex: number,
    itemId: string,
    delta: string,
    callId?: string,
  ): JsonObject {
    return this.emit(
      buildResponseFunctionCallArgumentsDeltaEvent(
        responseId,
        outputIndex,
        itemId,
        delta,
        callId,
      ),
    );
  }

  functionCallArgumentsDone(
    responseId: string,
    outputIndex: number,
    itemId: string,
    argumentsText: string,
    callId?: string,
  ): JsonObject {
    return this.emit(
      buildResponseFunctionCallArgumentsDoneEvent(
        responseId,
        outputIndex,
        itemId,
        argumentsText,
        callId,
      ),
    );
  }

  customToolInputDelta(
    responseId: string,
    outputIndex: number,
    itemId: string,
    delta: string,
    callId?: string,
  ): JsonObject {
    return this.emit(
      buildResponseCustomToolInputDeltaEvent(
        responseId,
        outputIndex,
        itemId,
        delta,
        callId,
      ),
    );
  }

  customToolInputDone(
    responseId: string,
    outputIndex: number,
    itemId: string,
    input: string,
    callId?: string,
  ): JsonObject {
    return this.emit(
      buildResponseCustomToolInputDoneEvent(
        responseId,
        outputIndex,
        itemId,
        input,
        callId,
      ),
    );
  }

  webSearchCallInProgress(
    responseId: string,
    outputIndex: number,
    itemIdOrItem?: string | JsonObject,
    item?: JsonObject,
  ): JsonObject {
    return this.emit(
      buildResponseWebSearchCallInProgressEvent(
        responseId,
        outputIndex,
        itemIdOrItem,
        item,
      ),
    );
  }

  webSearchCallSearching(
    responseId: string,
    outputIndex: number,
    itemIdOrItem?: string | JsonObject,
    item?: JsonObject,
  ): JsonObject {
    return this.emit(
      buildResponseWebSearchCallSearchingEvent(
        responseId,
        outputIndex,
        itemIdOrItem,
        item,
      ),
    );
  }

  webSearchCallCompleted(
    responseId: string,
    outputIndex: number,
    itemIdOrItem?: string | JsonObject,
    item?: JsonObject,
  ): JsonObject {
    return this.emit(
      buildResponseWebSearchCallCompletedEvent(
        responseId,
        outputIndex,
        itemIdOrItem,
        item,
      ),
    );
  }

  completed(response: JsonObject): JsonObject {
    this.assertOpen();
    const terminalResponse = withCompletedStatus(response);
    const event = this.emitEvent(buildResponseCompletedEvent(terminalResponse));
    this.terminal = {
      kind: "completed",
      response: cloneJsonValue(terminalResponse),
    };
    return event;
  }

  failed(response: JsonObject, failure: BridgeFailure): JsonObject {
    this.assertOpen();
    const terminalResponse = withFailedStatus(response, failure);
    const event = this.emitEvent(buildResponseFailedEvent(terminalResponse));
    this.terminal = {
      kind: "failed",
      response: cloneJsonValue(terminalResponse),
      failure,
    };
    return event;
  }

  incomplete(response: JsonObject, failure: BridgeFailure): JsonObject {
    this.assertOpen();
    const terminalResponse = withIncompleteStatus(response, failure);
    const event = this.emitEvent(
      buildResponseIncompleteEvent(terminalResponse),
    );
    this.terminal = {
      kind: "incomplete",
      response: cloneJsonValue(terminalResponse),
      failure,
    };
    return event;
  }

  clientAbort(): void {
    this.assertOpen();
    this.terminal = { kind: "cancelled" };
  }

  private emitEvent(event: JsonObject): JsonObject {
    const stamped = {
      ...cloneJsonValue(event),
      sequence_number: this.nextSequence,
    };
    this.sink(stamped);
    this.nextSequence += 1;
    return stamped;
  }

  private assertOpen(): void {
    if (this.terminal) {
      throw new Error(
        `Cannot emit a response event after terminal state "${this.terminal.kind}"`,
      );
    }
  }
}

function withCompletedStatus(response: JsonObject): JsonObject {
  const result = cloneJsonValue(response);
  result.status = "completed";
  result.error = null;
  result.incomplete_details = null;
  return result;
}

function withFailedStatus(
  response: JsonObject,
  failure: BridgeFailure,
): JsonObject {
  const result = cloneJsonValue(response);
  result.status = "failed";
  result.error = { code: failure.code, message: failure.message };
  result.incomplete_details = null;
  return result;
}

function withIncompleteStatus(
  response: JsonObject,
  failure: BridgeFailure,
): JsonObject {
  const result = cloneJsonValue(response);
  result.status = "incomplete";
  result.error = null;
  result.incomplete_details = { reason: failure.reason };
  return result;
}
