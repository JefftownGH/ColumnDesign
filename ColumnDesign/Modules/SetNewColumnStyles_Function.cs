/*namespace ColumnDesign.Modules
{
    public class SetNewColumnStyles_Function
    {
        Public Function BuildColumnStyles()
    'Create new dimension style and switch to it
    Dim NewDimStyle As AcadDimStyle
    Dim NewTextStyle As AcadTextStyle
    Dim iAltUnits As Integer
    Dim dDimScale As Double
    
    'Create new dim Style
    Set NewDimStyle = ThisDrawing.DimStyles.Add("VBA_Dimstyle")
    
    'Create text styles
    Set NewTextStyle = ThisDrawing.TextStyles.Add("VBA_txt_narrow")
    NewTextStyle.SetFont "Arial Narrow", False, False, 0, 0
    NewTextStyle.Height = 2
    
    Set NewTextStyle = ThisDrawing.TextStyles.Add("VBA_txt")
    NewTextStyle.SetFont "Arial", False, False, 0, 0
    NewTextStyle.Height = 2
    
    'Set newly created dim and text styles current
    ThisDrawing.ActiveDimStyle = NewDimStyle
    ThisDrawing.ActiveTextStyle = NewTextStyle
    
    'Alter the target "dimvar" values (these can be found in the dimstyle properties)
        'Lines Tab
        ThisDrawing.SetVariable "DIMCLRD", 151 'sets color for dimension lines
        ThisDrawing.SetVariable "DIMLTYPE", "ByBlock" 'sets dimension linetype
        ThisDrawing.SetVariable "DIMLWD", -2 'sets dimension lineweight, -2 is ByBlock
        ThisDrawing.SetVariable "DIMDLE", 1 'dim line extension past tick
        ThisDrawing.SetVariable "DIMDLI", 6 'default spacing between dimensions when using the DIMBASELINE command
        ThisDrawing.SetVariable "DIMCLRE", 150 'sets color for extension lines
        ThisDrawing.SetVariable "DIMLTEX1", "ByBlock" 'sets extension line 1 linetype
        ThisDrawing.SetVariable "DIMLTEX2", "ByBlock" 'sets extension line 2 linetype
        ThisDrawing.SetVariable "DIMLWE", -2 'sets extension line lineweight, -2 is by block
        ThisDrawing.SetVariable "DIMEXE", 1 'extension line extension beyond dim line
        ThisDrawing.SetVariable "DIMEXO", 1 'dim offset from origin
        
        'Symbols and Arrows Tab
        ThisDrawing.SetVariable "DIMSAH", 0 'use arrowhead blocks set by DIMBLK rather than DIMBLK1 and DIMBLK2
        ThisDrawing.SetVariable "DIMBLK", "_ARCHTICK" 'sets arrowhead block at dim lines
        ThisDrawing.SetVariable "DIMLDRBLK", "" 'sets arrowhead for leader line
        ThisDrawing.SetVariable "DIMASZ", 1.5 'arrowhead size
        
        'Text Tab
        ThisDrawing.SetVariable "DIMTXSTY", "VBA_txt" 'text style to use for this dimension style
        ThisDrawing.SetVariable "DIMCLRT", 150 'sets color for text
        ThisDrawing.SetVariable "DIMTXT", 2 'text height
        ThisDrawing.SetVariable "DIMTAD", 0 'horizontal text alignment, 0 is centered
        ThisDrawing.SetVariable "DIMJUST", 0 'vertical text alignment, 0 is centered
        ThisDrawing.SetVariable "DIMGAP", 1 'gap between text and dim lines
        ThisDrawing.SetVariable "DIMTIH", 0 'aligned with dimension line in conjunction with DIMTOH 'I think one is for dims inside the line and the other is for dims outside the line
        ThisDrawing.SetVariable "DIMTOH", 0 'aligned with dimension line in conjunctionw ith DIMTIH
        
        'Fit Tab
        ThisDrawing.SetVariable "DIMATFIT", 3 'for tight dimensions, moves text or arrows or both outside the extension lines
        ThisDrawing.SetVariable "DIMTMOVE", 0 '= 0, when you move text the dimension moves with it
        ThisDrawing.SetVariable "DIMSCALE", 1 'Controls dim size, recommendation: do not change!
        
        'Primary Units Tab
        ThisDrawing.SetVariable "DIMUNIT", 4 'Unit format, 1: scientific, 2: decimal, 3: engineering, 4: architectural, 5: fractional
        ThisDrawing.SetVariable "DIMTFAC", 1 'scales size of fractions
        ThisDrawing.SetVariable "DIMFRAC", 1 'fraction stacking format, 0: horizontal, 1: diagonal, 2: not stacked (1/2)
        ThisDrawing.SetVariable "DIMRND", 0.125 'Rounds dimension to this nearest value
        ThisDrawing.SetVariable "DIMZIN", 3 'Includes zero inches and suppresses zero feet

    'Copy new document dimvar settings into new dimstyle
    NewDimStyle.CopyFrom ThisDrawing
End Function

    }
}*/