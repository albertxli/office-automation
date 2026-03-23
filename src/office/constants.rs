//! Office COM constants — shape types, update options, calculation modes.

/// MsoShapeType values for identifying shape types.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MsoShapeType {
    Group = 6,
    LinkedOleObject = 10, // GOTCHA #2: NOT 7 (that's EmbeddedOleObject)
}

/// PpUpdateOption — how OLE link updates are handled.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PpUpdateOption {
    Manual = 1,    // GOTCHA #11: NOT 2
    Automatic = 2,
}

/// XlCalculation — Excel calculation mode.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlCalculation {
    Manual = -4135,
    Automatic = -4105,
}

/// MsoTriState — Office boolean tri-state.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MsoTriState {
    True = -1,
    False = 0,
}

/// PpAlertsLevel — PowerPoint alert display.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PpAlertsLevel {
    None = 0,
    All = 2,
}

/// MsoFillType values.
#[repr(i32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MsoFillType {
    Solid = 1,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_critical_constants() {
        // These constants were the source of bugs in the Python version
        assert_eq!(MsoShapeType::LinkedOleObject as i32, 10);
        assert_eq!(MsoShapeType::Group as i32, 6);
        assert_eq!(PpUpdateOption::Manual as i32, 1);
        assert_eq!(PpUpdateOption::Automatic as i32, 2);
        assert_eq!(XlCalculation::Manual as i32, -4135);
        assert_eq!(MsoTriState::True as i32, -1);
        assert_eq!(MsoTriState::False as i32, 0);
    }
}
