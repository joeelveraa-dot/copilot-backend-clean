// app/api/chat/route.ts
import { NextResponse } from "next/server";
import { COPILOT_FAE_SYSTEM_PROMPT, MODULE_PROMPTS } from "@/src/prompts/fae";

export async function POST(req: Request) {
  const incoming = await req.formData();

  const mode = (incoming.get("mode") || "chat").toString();

  // Inyecta identidad FAE + prompt del módulo como campos adicionales
  incoming.append("fae_context_system", COPILOT_FAE_SYSTEM_PROMPT);
  incoming.append("fae_context_module", MODULE_PROMPTS[mode] ?? MODULE_PROMPTS.chat);

  const url = process.env.BACKEND_INTERNAL_URL + "/chat";
  const res = await fetch(url, { method: "POST", body: incoming });

  const text = await res.text();
  if (!res.ok) return new NextResponse(text, { status: res.status });

  return new NextResponse(text, {
    status: 200,
    headers: { "Content-Type": "application/json" },
  });
}
