import { Route, Routes } from "react-router";
import Header from "./components/Header";
import Footer from "./components/Footer";
import Home from "./pages/Home";
import About from "./pages/About";
import WeeklyReports from "./pages/WeeklyReports";
import Research from "./pages/Research";
import Forms from "./pages/Forms";
import Contact from "./pages/Contact";

export default function App() {
  return (
    <div className="flex min-h-screen flex-col">
      <Header />
      <main className="flex-1">
        <Routes>
          <Route path="/" element={<Home />} />
          <Route path="/about" element={<About />} />
          <Route path="/weekly-reports" element={<WeeklyReports />} />
          <Route path="/research" element={<Research />} />
          <Route path="/forms" element={<Forms />} />
          <Route path="/contact" element={<Contact />} />
        </Routes>
      </main>
      <Footer />
    </div>
  );
}
