import React from "react";

interface GridbookIconProps {
  size?: number;
  className?: string;
  style?: React.CSSProperties;
}

/**
 * GridbookIcon: A clean, crisp vector icon representing a gridbook / spreadsheet notebook.
 * Features:
 * - Spiral binding rings on the left spine
 * - Vertical spine margin rule
 * - Spreadsheet table header row highlight
 * - Active spreadsheet cell highlight (Cell A2)
 * - Sharp grid matrix (2 columns x 3 rows)
 *
 * Inherits currentColor for stroke & fill accents, making it adapt
 * seamlessly to any container background or theme.
 */
export const GridbookIcon: React.FC<GridbookIconProps> = ({
  size = 16,
  className = "",
  style,
}) => {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      className={className}
      style={{ display: "inline-block", verticalAlign: "middle", ...style }}
      aria-hidden="true"
    >
      {/* Notebook outer cover */}
      <rect
        x="5"
        y="2.5"
        width="15.5"
        height="19"
        rx="2.5"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinejoin="round"
      />

      {/* Left spine margin rule */}
      <line
        x1="8.5"
        y1="2.5"
        x2="8.5"
        y2="21.5"
        stroke="currentColor"
        strokeWidth="1.2"
        opacity="0.45"
      />

      {/* Spiral binding rings */}
      <path
        d="M2.5 6.5H6.25M2.5 10.5H6.25M2.5 14.5H6.25M2.5 18.5H6.25"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinecap="round"
      />

      {/* Spreadsheet header row highlight */}
      <rect
        x="10"
        y="6"
        width="8"
        height="3.5"
        rx="0.5"
        fill="currentColor"
        opacity="0.22"
      />

      {/* Active spreadsheet cell highlight */}
      <rect
        x="10"
        y="9.5"
        width="4"
        height="4"
        rx="0.5"
        fill="currentColor"
        opacity="0.48"
      />

      {/* Grid columns */}
      <line
        x1="14"
        y1="6"
        x2="14"
        y2="17.5"
        stroke="currentColor"
        strokeWidth="1.2"
        opacity="0.85"
        strokeLinecap="round"
      />

      {/* Grid rows */}
      <line
        x1="10"
        y1="9.5"
        x2="18"
        y2="9.5"
        stroke="currentColor"
        strokeWidth="1.2"
        opacity="0.85"
        strokeLinecap="round"
      />
      <line
        x1="10"
        y1="13.5"
        x2="18"
        y2="13.5"
        stroke="currentColor"
        strokeWidth="1.2"
        opacity="0.85"
        strokeLinecap="round"
      />
    </svg>
  );
};

export default GridbookIcon;
