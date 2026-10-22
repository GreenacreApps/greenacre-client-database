// Test environment only: no schedule, storage access, or outbound email.
export default async () => new Response(JSON.stringify({ ok: true, disabled: true, environment: "isolated-test" }), { headers: { "Content-Type": "application/json" } });
