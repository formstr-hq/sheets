import { useState, type ReactNode } from "react";
import { BrowserRouter, matchPath, Navigate, Route, Routes, useLocation } from "react-router-dom";
import { Loader2 } from "lucide-react";
import { Landing } from "./components/Landing";
import { LoginModal, UnlockScreen } from "./components/LoginModal";
import { Dashboard } from "./components/Dashboard";
import { Editor } from "./components/Editor";
import { LogoMark } from "./components/Logo";
import { AuthProvider, useAuth } from "./providers/AuthProvider";
import { ProfileProvider } from "./providers/ProfileProvider";
import { LibraryProvider } from "./providers/LibraryProvider";
import { ToastProvider } from "./providers/ToastProvider";

function Splash() {
  return (
    <div className="splash">
      <LogoMark size={52} />
      <Loader2 size={18} className="spin" />
      <span>Opening Formstr Sheets…</span>
    </div>
  );
}

function AppRoutes() {
  const auth = useAuth();
  const location = useLocation();
  const [loginOpen, setLoginOpen] = useState(false);

  const sheetMatch = matchPath("/s/:naddr", location.pathname);
  const hasLinkKeys = location.hash.startsWith("#nkeys=");
  const canOpenLink = !!sheetMatch && hasLinkKeys;

  let content: ReactNode;

  if (!auth.ready) {
    content = <Splash />;
  } else if (!auth.pubkey) {
    content = canOpenLink ? (
      <Routes>
        <Route path="/s/:naddr" element={<Editor />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    ) : (
      <Landing onOpen={() => setLoginOpen(true)} />
    );
  } else if (auth.locked && !canOpenLink) {
    content = <UnlockScreen />;
  } else if (!auth.signer && !canOpenLink) {
    content = <Splash />;
  } else {
    content = (
      <Routes>
        <Route path="/" element={<Dashboard />} />
        <Route path="/s/:naddr" element={<Editor />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    );
  }

  return (
    <>
      {content}
      <LoginModal open={loginOpen} onClose={() => setLoginOpen(false)} />
    </>
  );
}

export default function App() {
  return (
    <BrowserRouter>
      <ToastProvider>
        <AuthProvider>
          <ProfileProvider>
            <LibraryProvider>
              <AppRoutes />
            </LibraryProvider>
          </ProfileProvider>
        </AuthProvider>
      </ToastProvider>
    </BrowserRouter>
  );
}
