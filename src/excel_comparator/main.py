"""Application entry point."""
import sys

sys.dont_write_bytecode = True

from excel_comparator.gui import ExcelComparatorApp


def run() -> None:
  """Start the desktop application."""
  app = ExcelComparatorApp()
  app.mainloop()


if __name__ == "__main__":
  run()
