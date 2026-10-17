import { useState } from "react";
import type React from "react";
import Popover from "@mui/material/Popover";
import VpnKeyOutlined from "@mui/icons-material/VpnKeyOutlined";

import styles from "./SheetIngestIdPopover.module.css";

export interface SheetIngestIdPopoverProps {
  /** Short label shown on the trigger chip. */
  label: string;
  /** IDs displayed in the popup. */
  values: readonly (string | number)[];
  /** Accessible label for the trigger. */
  ariaLabel?: string;
}

/** A compact count chip that opens a scrollable, package-styled list of IDs. */
export const SheetIngestIdPopover = ({
  label,
  values,
  ariaLabel,
}: SheetIngestIdPopoverProps): React.JSX.Element => {
  const [anchorEl, setAnchorEl] = useState<HTMLButtonElement | null>(null);
  const ids = values.map(String).map((id) => id.trim()).filter(Boolean);
  const open = Boolean(anchorEl);

  return (
    <>
      <button
        aria-label={ariaLabel ?? `${ids.length} ${label.toLowerCase()} IDs`}
        aria-expanded={open}
        aria-haspopup="dialog"
        className={styles.trigger}
        onClick={(event) => setAnchorEl((current) => current ? null : event.currentTarget)}
        type="button"
      >
        <VpnKeyOutlined aria-hidden="true" className={styles.triggerIcon} />
        <span>{label}</span>
        <span className={styles.count}>{ids.length}</span>
      </button>
      <Popover
        anchorEl={anchorEl}
        anchorOrigin={{ vertical: "bottom", horizontal: "left" }}
        classes={{ paper: styles.popoverPaper }}
        onClose={() => setAnchorEl(null)}
        open={open}
        sx={{ zIndex: 1500 }}
        transformOrigin={{ vertical: "top", horizontal: "left" }}
      >
        <section aria-label={`${label} IDs`} className={styles.popover}>
          <header className={styles.popoverHeader}>
            <div>
              <h3>{label} IDs</h3>
              <p>IDs included in this spreadsheet row</p>
            </div>
            <span className={styles.total}>{ids.length} total</span>
          </header>
          {ids.length > 0 ? (
            <div className={styles.values}>
              {ids.map((id, index) => (
                <span className={styles.value} key={`${id}-${index}`}>{id}</span>
              ))}
            </div>
          ) : (
            <p className={styles.empty}>No IDs provided.</p>
          )}
        </section>
      </Popover>
    </>
  );
};

export default SheetIngestIdPopover;
