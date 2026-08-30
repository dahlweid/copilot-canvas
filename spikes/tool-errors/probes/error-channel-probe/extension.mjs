// Which parts of a tool failure survive the crossing from an extension to the
// agent? See spikes/tool-errors/ for the question and the recorded answer.
//
// Five tools, one per candidate channel. Each carries a distinctive marker so a
// result can be attributed rather than recognised: if a marker is absent from
// what the agent receives, that channel dropped it.

import { joinSession, createCanvas, CanvasError } from "@github/copilot-sdk/extension";

const tool = (name, handler) => ({
    name,
    description: `Error-channel probe: ${name}. Call it and record verbatim what comes back.`,
    parameters: { type: "object", properties: {}, additionalProperties: false },
    handler,
});

// The other half of the question: a canvas action is a different transport from
// a tool, so what survives one says nothing about the other.
const probeCanvas = createCanvas({
    id: "error-channel-probe-canvas",
    displayName: "Error channel probe",
    description: "Measures what a canvas action failure carries to the agent.",
    open: async () => ({ title: "Error channel probe", status: "ready", url: "http://127.0.0.1:9/probe" }),
    actions: [
        {
            name: "action_throw_canvas_error",
            description: "Throws a CanvasError with a distinctive code and message.",
            inputSchema: { type: "object", properties: {}, additionalProperties: false },
            handler: async () => {
                throw new CanvasError("MARK-D-CODE", "MARK-D-MESSAGE");
            },
        },
        {
            name: "action_throw_plain",
            description: "Throws an ordinary Error.",
            inputSchema: { type: "object", properties: {}, additionalProperties: false },
            handler: async () => {
                throw new Error("MARK-E-MESSAGE");
            },
        },
        {
            name: "action_return_failure_object",
            description: "Returns a ToolResultObject-shaped failure.",
            inputSchema: { type: "object", properties: {}, additionalProperties: false },
            handler: async () => ({
                textResultForLlm: "MARK-F-TEXT",
                resultType: "failure",
            }),
        },
    ],
});

await joinSession({
    canvases: [probeCanvas],
    tools: [
        // A1: what office-canvas does today -- asToolError's shape, thrown.
        tool("probe_throw_coded", async () => {
            const err = new Error("MARK-A1-MESSAGE: file_locked: Another process is holding it.");
            err.code = "MARK-A1-CODE";
            err.data = { marker: "MARK-A1-DATA" };
            throw err;
        }),
        // A2: a bare throw, to separate "message dropped" from "shape rejected".
        tool("probe_throw_plain", async () => {
            throw new Error("MARK-A2-MESSAGE");
        }),
        // B: the SDK's documented ToolResultObject, flagged as a failure.
        tool("probe_return_failure_object", async () => ({
            textResultForLlm: "MARK-B-TEXT: file_locked: Another process is holding it.",
            resultType: "failure",
            error: "MARK-B-ERROR",
        })),
        // B2: same, without the `error` field -- wye() infers resultType from it,
        // so this separates the flag from the field.
        tool("probe_return_failure_no_error_field", async () => ({
            textResultForLlm: "MARK-B2-TEXT",
            resultType: "failure",
        })),
        // C: an ordinary object, which the SDK JSON.stringifies.
        tool("probe_return_plain_object", async () => ({
            code: "MARK-C-CODE",
            message: "MARK-C-MESSAGE",
        })),
    ],
});
