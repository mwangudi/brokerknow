import Flatpickr from "react-flatpickr";
import "flatpickr/dist/themes/light.css";

interface DatePickerProps {
  value: string; // ISO yyyy-mm-dd
  onChange: (value: string) => void;
  placeholder?: string;
  required?: boolean;
  /** Latest date allowed (default: today). Pass null to disable. */
  maxDate?: string | Date | null;
  /** Earliest date allowed. */
  minDate?: string | Date | null;
  ariaLabel?: string;
  id?: string;
}

const inputClass =
  "w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm text-gray-700 shadow-theme-xs placeholder:text-gray-400 focus:border-brand-300 focus:outline-none focus:ring-4 focus:ring-brand-500/10 dark:border-gray-700 dark:bg-gray-800 dark:text-gray-400 dark:placeholder:text-gray-500";

/**
 * Themed date picker (Flatpickr) returning ISO yyyy-mm-dd strings,
 * compatible with the rest of the app's input styling.
 */
export default function DatePicker({
  value,
  onChange,
  placeholder = "Select a date",
  required = false,
  maxDate = "today",
  minDate,
  ariaLabel,
  id,
}: DatePickerProps) {
  return (
    <Flatpickr
      value={value || undefined}
      onChange={(dates) => {
        if (!dates.length) {
          onChange("");
          return;
        }
        const d = dates[0];
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        const dd = String(d.getDate()).padStart(2, "0");
        onChange(`${yyyy}-${mm}-${dd}`);
      }}
      options={{
        dateFormat: "Y-m-d",
        altInput: true,
        altFormat: "d M Y",
        allowInput: true,
        maxDate: maxDate ?? undefined,
        minDate: minDate ?? undefined,
      }}
      placeholder={placeholder}
      className={inputClass}
      required={required}
      aria-label={ariaLabel}
      id={id}
    />
  );
}
