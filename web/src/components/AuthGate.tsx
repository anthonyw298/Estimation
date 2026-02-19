'use client';

import { useState, useEffect, useCallback } from 'react';
import { Lock } from 'lucide-react';

const CORRECT_PASSWORD = 'password';
const AUTH_KEY = 'ugv_authenticated';

export default function AuthGate({ children }: { children: React.ReactNode }) {
  const [authenticated, setAuthenticated] = useState(false);
  const [checking, setChecking] = useState(true);
  const [password, setPassword] = useState('');
  const [error, setError] = useState(false);

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
        setPassword('');
      }
    },
    [password],
  );

  // Avoid flash while checking sessionStorage
  if (checking) {
    return (
      <div className="min-h-screen bg-[#08080e] flex items-center justify-center">
        <div className="w-5 h-5 border-2 border-blue-500 border-t-transparent rounded-full animate-spin" />
      </div>
    );
  }

  if (authenticated) {
    return <>{children}</>;
  }

  return (
    <div className="min-h-screen bg-[#08080e] flex items-center justify-center px-4">
      <div className="w-full max-w-sm">
        {/* Logo / Brand */}
        <div className="text-center mb-8">
          <div className="inline-flex items-center justify-center w-14 h-14 rounded-2xl bg-[#111118] border border-[#1e1e2a] mb-4 shadow-lg shadow-black/20">
            <Lock className="w-6 h-6 text-blue-400" />
          </div>
          <h1 className="text-xl font-bold text-[#eeeef2] tracking-tight">
            United Glass Ventures
          </h1>
          <p className="text-sm text-[#55566a] mt-1">Estimator</p>
        </div>

        {/* Password Form */}
        <form
          onSubmit={handleSubmit}
          className="bg-[#111118] border border-[#1e1e2a] rounded-2xl p-6 space-y-4 shadow-2xl shadow-black/30"
        >
          <div>
            <label className="block text-sm font-medium text-[#8b8d9a] mb-1.5">
              Password
            </label>
            <input
              type="password"
              autoFocus
              value={password}
              onChange={(e) => {
                setPassword(e.target.value);
                if (error) setError(false);
              }}
              placeholder="Enter password"
              className="w-full bg-[#0c0c12] border border-[#1e1e2a] text-white rounded-lg px-3 py-2.5 text-sm placeholder:text-[#3e3f4d] focus:outline-none focus:ring-2 focus:ring-blue-500/40 focus:border-blue-500/40 transition-all duration-200"
            />
          </div>

          {error && (
            <p className="text-xs text-red-400 font-medium">
              Incorrect password. Please try again.
            </p>
          )}

          <button
            type="submit"
            className="w-full py-2.5 text-sm font-medium bg-blue-600 hover:bg-blue-500 active:scale-[0.98] text-white rounded-lg transition-all duration-200 shadow-md shadow-blue-500/10"
          >
            Sign In
          </button>
        </form>
      </div>
    </div>
  );
}
