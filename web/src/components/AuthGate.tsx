'use client';

import { useState, useEffect, useCallback } from 'react';
import { Lock, Building2, Sparkles } from 'lucide-react';

const CORRECT_PASSWORD = 'UnitedGlass01!#';
const AUTH_KEY = 'ugv_authenticated';

export default function AuthGate({ children }: { children: React.ReactNode }) {
  const [authenticated, setAuthenticated] = useState(false);
  const [checking, setChecking] = useState(true);
  const [password, setPassword] = useState('');
  const [error, setError] = useState(false);
  const [shaking, setShaking] = useState(false);

  // Check sessionStorage on mount
  useEffect(() => {
    if (typeof window !== 'undefined') {
      const stored = sessionStorage.getItem(AUTH_KEY);
      if (stored === 'true') {
        setAuthenticated(true);
      }
    }
    setChecking(false);
  }, []);

  const handleSubmit = useCallback(
    (e: React.FormEvent) => {
      e.preventDefault();
      if (password === CORRECT_PASSWORD) {
        sessionStorage.setItem(AUTH_KEY, 'true');
        setAuthenticated(true);
        setError(false);
      } else {
        setError(true);
        setShaking(true);
        setPassword('');
        setTimeout(() => setShaking(false), 500);
      }
    },
    [password],
  );

  // Avoid flash while checking sessionStorage
  if (checking) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="w-6 h-6 border-2 border-blue-500/30 border-t-blue-500 rounded-full animate-spin" />
      </div>
    );
  }

  if (authenticated) {
    return <>{children}</>;
  }

  return (
    <div className="min-h-screen flex items-center justify-center px-4 relative">

      <div className="w-full max-w-[380px] relative z-10 animate-fade-up opacity-0">
        {/* Logo / Brand */}
        <div className="text-center mb-10">
          <div className="relative inline-flex items-center justify-center mb-6">
            {/* Glow ring behind icon */}
            <div className="absolute inset-0 w-20 h-20 rounded-2xl bg-blue-500/10 blur-xl animate-breathe" />
            <div className="animated-border w-20 h-20 rounded-2xl bg-gradient-to-br from-[#111118] to-[#0c0c14] flex items-center justify-center border border-[#1e1e2a] shadow-2xl shadow-blue-500/10 relative">
              <Building2 className="w-9 h-9 text-blue-400" />
            </div>
          </div>
          <h1 className="text-3xl font-bold gradient-text-static tracking-tight">
            United Glass Ventures
          </h1>
          <p className="text-sm text-[#ffffff] mt-2 font-medium tracking-[0.2em] uppercase flex items-center justify-center gap-1.5">
            <Sparkles className="w-3.5 h-3.5 text-[#3b82f6]/50" />
            Estimator Pro
            <Sparkles className="w-3.5 h-3.5 text-[#8b5cf6]/50" />
          </p>
        </div>

        {/* Password Form */}
        <form
          onSubmit={handleSubmit}
          className={`card-glow glass-card rounded-2xl p-8 space-y-6 shadow-2xl shadow-black/50 ${shaking ? 'animate-shake' : ''}`}
        >
          <div>
            <label className="block text-sm font-medium text-[#ffffff] mb-2.5">
              Password
            </label>
            <div className="relative group">
              <div className="absolute left-4 top-1/2 -translate-y-1/2 text-[#ffffff] group-focus-within:text-[#3b82f6] transition-colors duration-200">
                <Lock className="w-4 h-4" />
              </div>
              <input
                type="password"
                autoFocus
                value={password}
                onChange={(e) => {
                  setPassword(e.target.value);
                  if (error) setError(false);
                }}
                placeholder="Enter password"
                className="w-full bg-[#0c0c12] border border-[#1e1e2a] text-white rounded-xl pl-11 pr-4 py-3.5 text-sm placeholder:text-[#ffffff] focus:outline-none focus:ring-2 focus:ring-blue-500/30 focus:border-blue-500/40 transition-colors duration-200"
              />
            </div>
          </div>

          {error && (
            <div className="flex items-center gap-2 px-3 py-2.5 bg-red-500/5 border border-red-500/10 rounded-lg animate-fade-in">
              <span className="w-1.5 h-1.5 rounded-full bg-red-400 animate-subtle-pulse" />
              <p className="text-xs text-red-400 font-medium">
                Incorrect password. Please try again.
              </p>
            </div>
          )}

          <button
            type="submit"
            className="w-full py-3.5 text-sm font-semibold bg-gradient-to-r from-blue-600 via-blue-500 to-indigo-500 hover:brightness-110 text-white rounded-xl transition-colors duration-200"
          >
            Sign In
          </button>
        </form>

        {/* Subtle footer */}
        <div className="text-center mt-8">
          <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-[#111118]/50 border border-[#1e1e2a]/50">
            <div className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-subtle-pulse" />
            <p className="text-[10px] text-[#ffffff] tracking-wide uppercase font-medium">
              Secure Session
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}
