import React, { useState } from 'react';
import { Building2, Lock, User, Eye, EyeOff, ArrowLeft, AlertCircle } from 'lucide-react';

export interface UserSession {
  id: number;
  username: string;
  displayName: string;
  role: string;
  loginTime: string;
}

interface LoginViewProps {
  onLoginSuccess: (user: UserSession) => void;
}

export default function LoginView({ onLoginSuccess }: LoginViewProps) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [rememberMe, setRememberMe] = useState(true);
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!username.trim() || !password.trim()) {
      setErrorMsg('يرجى إدخال اسم المستخدم وكلمة المرور');
      return;
    }

    setLoading(true);
    setErrorMsg('');

    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          username: username.trim(),
          password: password.trim()
        })
      });

      const data = await res.json();

      if (!res.ok || !data.success) {
        throw new Error(data.message || 'اسم المستخدم أو كلمة المرور غير صحيحة');
      }

      if (rememberMe) {
        localStorage.setItem('alroyal_session', JSON.stringify(data.user));
      } else {
        sessionStorage.setItem('alroyal_session', JSON.stringify(data.user));
      }

      onLoginSuccess(data.user);
    } catch (err: any) {
      setErrorMsg(err.message || 'تعذر الاتصال بخادم النظام');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div
      className="min-h-screen w-full relative flex items-center justify-center p-4 bg-gradient-to-b from-white via-slate-50 to-amber-50/30 text-slate-800 font-sans select-none overflow-hidden"
      dir="rtl"
    >
      {/* Background Animated Ambient Lights & Shapes */}
      <div className="absolute inset-0 pointer-events-none overflow-hidden">
        {/* Floating Warm Champagne Orb 1 */}
        <div className="absolute -top-24 right-1/4 w-[500px] h-[500px] rounded-full bg-amber-200/35 blur-[120px] animate-float-slow animate-pulse-glow" />
        
        {/* Floating Warm Champagne Orb 2 */}
        <div className="absolute -bottom-24 left-1/4 w-[520px] h-[520px] rounded-full bg-amber-100/50 blur-[130px] animate-float-reverse" />
        
        {/* Architectural Grid Pattern */}
        <div 
          className="absolute inset-0 opacity-[0.035] bg-[linear-gradient(to_right,#0f172a_1px,transparent_1px),linear-gradient(to_bottom,#0f172a_1px,transparent_1px)] bg-[size:4rem_4rem]" 
        />

        {/* Ambient Top Subtle Golden Line */}
        <div className="absolute top-0 inset-x-0 h-px bg-gradient-to-r from-transparent via-amber-400/40 to-transparent" />
      </div>

      {/* Main Login Card */}
      <div className="relative w-full max-w-md z-10 animate-card-entrance">
        <div className="bg-white/95 backdrop-blur-xl rounded-3xl p-8 sm:p-10 shadow-[0_20px_50px_-15px_rgba(15,23,42,0.08),0_1px_3px_rgba(15,23,42,0.04)] border border-slate-200/90 relative">
          
          {/* Subtle top card golden accent */}
          <div className="absolute top-0 inset-x-12 h-[2px] bg-gradient-to-r from-transparent via-amber-400/60 to-transparent rounded-full" />

          {/* Header & Logo */}
          <div className="text-center mb-8">
            <div className="w-16 h-16 mx-auto mb-4 rounded-2xl bg-gradient-to-tr from-amber-500 via-amber-400 to-amber-500 flex items-center justify-center text-slate-950 shadow-md shadow-amber-500/25 animate-float-slow">
              <Building2 className="w-8 h-8" />
            </div>

            <h1 className="text-2xl font-black tracking-tight text-slate-900">
              الرويال للتطوير العقاري
            </h1>
            <p className="text-xs text-amber-600 font-bold mt-1.5 tracking-wide">
              النظام المالي والمحاسبي
            </p>
          </div>

          {/* Error Alert */}
          {errorMsg && (
            <div className="mb-5 p-3 rounded-xl bg-red-50 border border-red-200 text-red-700 text-xs flex items-center gap-2">
              <AlertCircle className="w-4 h-4 shrink-0 text-red-600" />
              <span className="font-medium">{errorMsg}</span>
            </div>
          )}

          {/* Form */}
          <form onSubmit={handleSubmit} className="space-y-4">
            <div>
              <label className="block text-xs font-bold text-slate-700 mb-1.5">
                اسم المستخدم
              </label>
              <div className="relative flex items-center">
                <div className="absolute right-3.5 text-slate-400 pointer-events-none">
                  <User className="w-4 h-4" />
                </div>
                <input
                  type="text"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  placeholder="اسم المستخدم"
                  required
                  autoFocus
                  className="w-full bg-slate-50/80 border border-slate-200/90 focus:bg-white focus:border-amber-500 rounded-xl pr-10 pl-4 py-3 text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-amber-500/15 transition font-medium"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold text-slate-700 mb-1.5">
                كلمة المرور
              </label>
              <div className="relative flex items-center">
                <div className="absolute right-3.5 text-slate-400 pointer-events-none">
                  <Lock className="w-4 h-4" />
                </div>
                <input
                  type={showPassword ? 'text' : 'password'}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="كلمة المرور"
                  required
                  className="w-full bg-slate-50/80 border border-slate-200/90 focus:bg-white focus:border-amber-500 rounded-xl pr-10 pl-11 py-3 text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-amber-500/15 transition font-mono"
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(!showPassword)}
                  className="absolute left-3.5 text-slate-400 hover:text-amber-600 transition cursor-pointer"
                  tabIndex={-1}
                >
                  {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                </button>
              </div>
            </div>

            <div className="pt-1">
              <label className="flex items-center gap-2 text-xs text-slate-600 hover:text-slate-800 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={rememberMe}
                  onChange={(e) => setRememberMe(e.target.checked)}
                  className="w-4 h-4 rounded border-slate-300 accent-amber-600 cursor-pointer"
                />
                <span>تذكر تسجيل الدخول</span>
              </label>
            </div>

            <button
              type="submit"
              disabled={loading}
              className="w-full mt-2 py-3 px-4 rounded-xl bg-gradient-to-r from-amber-500 via-amber-600 to-amber-500 hover:from-amber-600 hover:to-amber-700 text-white font-bold text-sm shadow-md shadow-amber-500/25 active:scale-[0.99] transition flex items-center justify-center gap-2 disabled:opacity-60 cursor-pointer"
            >
              {loading ? (
                <span>جاري تسجيل الدخول...</span>
              ) : (
                <>
                  <span>تسجيل الدخول</span>
                  <ArrowLeft className="w-4 h-4" />
                </>
              )}
            </button>
          </form>
        </div>

        {/* Minimal Footer */}
        <div className="text-center mt-5 text-slate-400 text-xs">
          الرويال للتطوير العقاري © {new Date().getFullYear()}
        </div>
      </div>
    </div>
  );
}
