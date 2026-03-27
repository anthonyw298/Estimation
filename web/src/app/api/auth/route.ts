import { NextResponse } from 'next/server';

const attempts = new Map<string, { count: number; resetAt: number }>();
const MAX_ATTEMPTS = 5;
const WINDOW_MS = 15 * 60 * 1000; // 15 minutes

function getClientIp(request: Request): string {
  return request.headers.get('x-forwarded-for')?.split(',')[0]?.trim() || 'unknown';
}

export async function POST(request: Request) {
  const correct = process.env.AUTH_PASSWORD;
  if (!correct) {
    return NextResponse.json({ ok: false, error: 'Auth not configured' }, { status: 503 });
  }

  // Rate limiting
  const ip = getClientIp(request);
  const now = Date.now();
  const record = attempts.get(ip);
  if (record && now < record.resetAt && record.count >= MAX_ATTEMPTS) {
    return NextResponse.json({ ok: false, error: 'Too many attempts' }, { status: 429 });
  }

  const { password } = await request.json();

  if (password === correct) {
    attempts.delete(ip);
    return NextResponse.json({ ok: true });
  }

  // Track failed attempt
  if (!record || now >= record.resetAt) {
    attempts.set(ip, { count: 1, resetAt: now + WINDOW_MS });
  } else {
    record.count++;
  }

  return NextResponse.json({ ok: false }, { status: 401 });
}
