import { NextResponse } from "next/server";

const USERNAME = "admin";
const PASSWORD = "admin";
const AUTH_COOKIE = "ocr_auth";

export async function POST(req: Request) {
  try {
    const { username, password } = await req.json();

    if (username !== USERNAME || password !== PASSWORD) {
      return NextResponse.json({ error: "Invalid credentials" }, { status: 401 });
    }

    const res = NextResponse.json({ ok: true });
    res.cookies.set({
      name: AUTH_COOKIE,
      value: "1",
      httpOnly: true,
      sameSite: "lax",
      secure: process.env.NODE_ENV === "production",
      path: "/",
    });
    return res;
  } catch (error) {
    return NextResponse.json(
      { error: "Invalid request", details: error instanceof Error ? error.message : String(error) },
      { status: 400 }
    );
  }
}
