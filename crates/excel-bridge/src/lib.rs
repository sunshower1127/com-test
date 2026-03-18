use com_core::{DispatchObject, Variant};
use windows::core::Result;

// ============================================================
// ExcelApp
// ============================================================

pub struct ExcelApp {
    disp: DispatchObject,
}

impl ExcelApp {
    /// Launch a new Excel instance
    pub fn new() -> Result<Self> {
        let disp = DispatchObject::create("Excel.Application")?;
        Ok(Self { disp })
    }

    /// Show or hide the Excel window
    pub fn set_visible(&self, visible: bool) -> Result<()> {
        self.disp.put("Visible", visible)
    }

    /// Enable or disable alert dialogs (save confirmation, etc.)
    pub fn set_display_alerts(&self, enable: bool) -> Result<()> {
        self.disp.put("DisplayAlerts", enable)
    }

    /// Access the Workbooks collection
    pub fn workbooks(&self) -> Result<Workbooks> {
        let disp = self.disp.get("Workbooks")?.into_dispatch()?;
        Ok(Workbooks { disp })
    }

    /// Quit Excel
    pub fn quit(&self) -> Result<()> {
        self.disp.call("Quit", &[])?;
        Ok(())
    }
}

impl Drop for ExcelApp {
    fn drop(&mut self) {
        let _ = self.disp.put("DisplayAlerts", false);
        let _ = self.disp.call("Quit", &[]);
    }
}

// ============================================================
// Workbooks
// ============================================================

pub struct Workbooks {
    disp: DispatchObject,
}

impl Workbooks {
    /// Create a new empty workbook
    pub fn add(&self) -> Result<Workbook> {
        let disp = self.disp.call("Add", &[])?.into_dispatch()?;
        Ok(Workbook { disp })
    }

    /// Open an existing workbook file
    pub fn open(&self, path: &str) -> Result<Workbook> {
        let disp = self
            .disp
            .call("Open", &[Variant::from(path)])?
            .into_dispatch()?;
        Ok(Workbook { disp })
    }
}

// ============================================================
// Workbook
// ============================================================

pub struct Workbook {
    disp: DispatchObject,
}

impl Workbook {
    /// Get a worksheet by 1-based index
    pub fn sheet(&self, index: i32) -> Result<Worksheet> {
        let sheets = self.disp.get("Sheets")?.into_dispatch()?;
        let disp = sheets
            .get_by("Item", &[Variant::I32(index)])?
            .into_dispatch()?;
        Ok(Worksheet { disp })
    }

    /// Save the workbook (must have been saved before, or use save_as)
    pub fn save(&self) -> Result<()> {
        self.disp.call("Save", &[])?;
        Ok(())
    }

    /// Save the workbook to a specific path
    pub fn save_as(&self, path: &str) -> Result<()> {
        self.disp.call("SaveAs", &[Variant::from(path)])?;
        Ok(())
    }

    /// Close the workbook without saving
    pub fn close(&self) -> Result<()> {
        self.disp
            .call("Close", &[Variant::Bool(false)])?;
        Ok(())
    }
}

// ============================================================
// Worksheet
// ============================================================

pub struct Worksheet {
    disp: DispatchObject,
}

impl Worksheet {
    /// Get worksheet name
    pub fn name(&self) -> Result<String> {
        let v = self.disp.get("Name")?;
        Ok(v.as_str().unwrap_or("").to_string())
    }

    /// Get a cell by 1-based row and column
    pub fn cell(&self, row: i32, col: i32) -> Result<Range> {
        let disp = self
            .disp
            .get_by("Cells", &[Variant::I32(row), Variant::I32(col)])?
            .into_dispatch()?;
        Ok(Range { disp })
    }

    /// Get a range by A1-notation (e.g. "A1:B5")
    pub fn range(&self, address: &str) -> Result<Range> {
        let disp = self
            .disp
            .get_by("Range", &[Variant::from(address)])?
            .into_dispatch()?;
        Ok(Range { disp })
    }
}

// ============================================================
// Range (cell or range of cells)
// ============================================================

pub struct Range {
    disp: DispatchObject,
}

impl Range {
    /// Get the cell value
    pub fn value(&self) -> Result<Variant> {
        self.disp.get("Value")
    }

    /// Set the cell value
    pub fn set_value(&self, value: impl Into<Variant>) -> Result<()> {
        self.disp.put("Value", value)
    }

    /// Set a formula (e.g. "=A1+B1")
    pub fn set_formula(&self, formula: &str) -> Result<()> {
        self.disp.put("Formula", formula)
    }

    /// Get the formula string
    pub fn formula(&self) -> Result<String> {
        let v = self.disp.get("Formula")?;
        Ok(v.as_str().unwrap_or("").to_string())
    }
}
