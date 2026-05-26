import { Routes, Route } from "react-router";
import AppLayout from "./layout/AppLayout";
import { ScrollToTop } from "./components/common/ScrollToTop";
import LoginPage from "./pages/LoginPage";
import RegisterPage from "./pages/RegisterPage";
import DashboardPage from "./pages/DashboardPage";
import ProfilePage from "./pages/ProfilePage";
import StatementPage from "./pages/StatementPage";
import OrdersPage from "./pages/OrdersPage";
import MarketPricesPage from "./pages/MarketPricesPage";
import RequestPaymentPage from "./pages/RequestPaymentPage";
import ChangePasswordPage from "./pages/ChangePasswordPage";
import { Navigate } from "react-router";
import { useAuth } from "./context/AuthContext";

function ProtectedRoutes() {
  const { user, token } = useAuth();
  if (!token || !user) return <Navigate to="/login" replace />;
  // Force the password change before any other authenticated page is rendered.
  if (user.mustChangePassword) return <Navigate to="/change-password" replace />;
  return <AppLayout />;
}

export default function App() {
  return (
    <>
      <ScrollToTop />
      <Routes>
        <Route path="/login" element={<LoginPage />} />
        <Route path="/register" element={<RegisterPage />} />
        <Route path="/change-password" element={<ChangePasswordPage />} />
        <Route element={<ProtectedRoutes />}>
          <Route index element={<DashboardPage />} />
          <Route path="profile" element={<ProfilePage />} />
          <Route path="statement" element={<StatementPage />} />
          <Route path="orders" element={<OrdersPage />} />
          <Route path="market-prices" element={<MarketPricesPage />} />
          <Route path="request-payment" element={<RequestPaymentPage />} />
        </Route>
      </Routes>
    </>
  );
}
