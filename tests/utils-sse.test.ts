import { describe, expect, test } from "bun:test";
import { readSseEvents } from "../src/proxy/utils";

describe("readSseEvents", () => {
  test("flushes a final event without a trailing newline", async () => {
    const events = await collectEvents(["event: result\ndata: final"]);

    expect(events).toEqual([{ event: "result", data: "final" }]);
  });

  test("removes at most one optional space after the data colon", async () => {
    const events = await collectEvents(["data:   preserve\n\n"]);

    expect(events).toEqual([{ event: "message", data: "  preserve" }]);
  });

  test("decodes UTF-8 split across byte chunks", async () => {
    const bytes = new TextEncoder().encode("data: räksmörgås\n\n");
    const split = bytes.indexOf(0xc3) + 1;
    const events = await collectByteEvents([
      bytes.slice(0, split),
      bytes.slice(split),
    ]);

    expect(events).toEqual([
      { event: "message", data: "räksmörgås" },
    ]);
  });

  test("accepts CR, LF, and CRLF line endings across chunks", async () => {
    const events = await collectEvents([
      "event: first\rdata: one\r\r",
      "event: second\r",
      "\ndata: two\n\n",
    ]);

    expect(events).toEqual([
      { event: "first", data: "one" },
      { event: "second", data: "two" },
    ]);
  });
});

async function collectEvents(chunks: string[]) {
  return collectByteEvents(chunks.map((chunk) => new TextEncoder().encode(chunk)));
}

async function collectByteEvents(chunks: Uint8Array[]) {
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      for (const chunk of chunks) controller.enqueue(chunk);
      controller.close();
    },
  });
  const events: Array<{ event: string; data: string }> = [];
  for await (const event of readSseEvents(stream)) events.push(event);
  return events;
}
