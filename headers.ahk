#Requires AutoHotkey v2.0

#SingleInstance Force

; Declare GUI and variables

stageList := ['First time Install',
    'Reselect',
    'Add on',
    'Completion',
    'Shortage',
    'Inspection',
    'Repair',
    'Refurb',
    'Demo',
    'Reinstall',
    'Prep',
    'Adjusted Labor']

commodityList := ['(Vinyl)', 'Laminate', 'LVP Click', 'Roll Vinyl', 'Rubber', 'VP Glue', '(Wood)', 'Gluedown wood',
    '(Carpet)', 'Carpet', 'Carpet Square', '(Misc)', 'House file', 'Slab', 'Multiple', '(Tile)', 'Multiple',
    'Tile Floor', 'Tile Wall', 'Tile Fireplace', 'Tile Backsplash', 'Tile Balcony', 'Tile Patio', 'Tile Pavers',
    'Grout', 'Caulking', '(Separated)', 'Waterproof', 'Tile Wall Master', 'Tile Wall Secondary',
    'Tile Wainscot', 'Tile Tub Area', 'Tile Seat', 'Tile Mudpan', 'Tile Dog Wash']

prefixList := ["NONE", "REORDER", "MODEL", "VIP", "OUT OF WARRANTY", "OCCUPIED", "CASH BEFORE DELIVERY", "THIRD PARTY"]

selectedStage := "", selectedCommodity := ""

; Main GUI

myGui := Gui()

myGui.Add("Text", , "Stage:")

cbStage := myGui.Add("DropDownList", "w250 Choose1", stageList)

cbStage.OnEvent("Change", (*) => UpdateResult())

myGui.Add("Text", , "Commodity:")

cbCommodity := myGui.Add("DropDownList", "w250 Choose1", commodityList)

cbCommodity.OnEvent("Change", (*) => UpdateResult())

myGui.Add("Text", , "Prefix:")

cbPrefix := myGui.Add("DropDownList", "w250 Choose1", prefixList)

cbPrefix.OnEvent("Change", (*) => UpdateResult())

cbOverride := myGui.Add("Checkbox", "Checked0", "Override Mode")

cbOverride.OnEvent("Click", (*) => UpdateResult())

myGui.Add("Text", , "Category:")

cbCategory := myGui.Add("DropDownList", "w250 Choose1", ["ACCLIMATION", "ACID WASH", "GROUT STAIN", "MATERIAL DELIVERY",
    "MATERIAL PICK UP", "MATERIAL NOT RETURNED", "DRAW ONLY", "TRIP CHARGE", "MOVERS", "REIMBURSEMENT", "HO SETTLEMENT"])

cbCategory.OnEvent("Change", (*) => UpdateResult())

myGui.Add("Text", , "Material:")

cbMaterial := myGui.Add("DropDownList", "w250 Choose1", ["TILE", "VINYL", "WOOD", "CARPET"])

cbMaterial.OnEvent("Change", (*) => UpdateResult())

tbResult := myGui.Add("Edit", "w400 h60 ReadOnly")

myGui.Add("Button", , "Copy to Clipboard").OnEvent("Click", (*) => A_Clipboard := tbResult.Value)

cbAlwaysOnTop := myGui.Add("Checkbox", , "Always on Top")

cbAlwaysOnTop.OnEvent("Click", (*) => ToggleAlwaysOnTop())

myGui.Title := "Stage + Commodity Lookup Tool"

myGui.Show()

ToggleAlwaysOnTop() {

    myGui.Opt(cbAlwaysOnTop.Value ? "+AlwaysOnTop" : "-AlwaysOnTop")

}

; --- Lookup Table (Auto-Generated from Excel) ---

prefixMap := Map()

prefixMap["NONE"] := ""
prefixMap["REORDER"] := "RO"
prefixMap["MODEL"] := "MODEL"
prefixMap["VIP"] := "VIP"
prefixMap["OUT OF WARRANTY"] := "OOW"
prefixMap["OCCUPIED"] := "OCC"
prefixMap["CASH BEFORE DELIVERY"] := "CBD"
prefixMap["THIRD PARTY"] := "3PTY"

lookup := Map()

lookup["First time Install|(Vinyl)"] := ""
lookup["Reselect|(Vinyl)"] := ""
lookup["Add on|(Vinyl)"] := ""
lookup["Completion|(Vinyl)"] := ""
lookup["Shortage|(Vinyl)"] := ""
lookup["Inspection|(Vinyl)"] := ""
lookup["Repair|(Vinyl)"] := ""
lookup["Refurb|(Vinyl)"] := ""
lookup["Demo|(Vinyl)"] := ""
lookup["Reinstall|(Vinyl)"] := ""
lookup["Prep|(Vinyl)"] := ""
lookup["Adjusted Labor|(Vinyl)"] := ""
lookup["First time Install|Laminate"] := "LAMINATE"
lookup["Reselect|Laminate"] := "LAMINATE RESELECT"
lookup["Add on|Laminate"] := "LAMINATE ADD ON"
lookup["Completion|Laminate"] := "LAMINATE COMP"
lookup["Shortage|Laminate"] := "LAMINATE SHORT"
lookup["Inspection|Laminate"] := "LAMINATE INSP"
lookup["Repair|Laminate"] := "LAMINATE REP"
lookup["Refurb|Laminate"] := "LAMINATE REFURB"
lookup["Demo|Laminate"] := "LAMINATE DEMO"
lookup["Reinstall|Laminate"] := "LAMINATE REINSTALL"
lookup["Prep|Laminate"] := "LAMINATE PREP"
lookup["Adjusted Labor|Laminate"] := "LAMINATE ADJ LABOR"
lookup["First time Install|LVP Click"] := "LVP CLICK"
lookup["Reselect|LVP Click"] := "LVP CLICK RESELECT"
lookup["Add on|LVP Click"] := "LVP CLICK ADD ON"
lookup["Completion|LVP Click"] := "LVP CLICK COMP"
lookup["Shortage|LVP Click"] := "LVP CLICK SHORT"
lookup["Inspection|LVP Click"] := "LVP CLICK INSP"
lookup["Repair|LVP Click"] := "LVP CLICK REP"
lookup["Refurb|LVP Click"] := "LVP CLICK REFURB"
lookup["Demo|LVP Click"] := "LVP CLICK DEMO"
lookup["Reinstall|LVP Click"] := "LVP CLICK REINSTALL"
lookup["Prep|LVP Click"] := "LVP CLICK PREP"
lookup["Adjusted Labor|LVP Click"] := "LVP CLICK ADJ LABOR"
lookup["First time Install|Roll Vinyl"] := "ROLL VINYL"
lookup["Reselect|Roll Vinyl"] := "ROLL VINYL RESELECT"
lookup["Add on|Roll Vinyl"] := "ROLL VINYL ADD ON"
lookup["Completion|Roll Vinyl"] := "ROLL VINYL COMP"
lookup["Shortage|Roll Vinyl"] := "ROLL VINYL SHORT"
lookup["Inspection|Roll Vinyl"] := "ROLL VINYL INSP"
lookup["Repair|Roll Vinyl"] := "ROLL VINYL REP"
lookup["Refurb|Roll Vinyl"] := "ROLL VINYL REFURB"
lookup["Demo|Roll Vinyl"] := "ROLL VINYL DEMO"
lookup["Reinstall|Roll Vinyl"] := "ROLL VINYL REINSTALL"
lookup["Prep|Roll Vinyl"] := "ROLL VINYL PREP"
lookup["Adjusted Labor|Roll Vinyl"] := "ROLL VINYL ADJ LABOR"
lookup["First time Install|Rubber"] := "RUBBER"
lookup["Reselect|Rubber"] := "RUBBER RESELECT"
lookup["Add on|Rubber"] := "RUBBER ADD ON"
lookup["Completion|Rubber"] := "RUBBER COMP"
lookup["Shortage|Rubber"] := "RUBBER SHORT"
lookup["Inspection|Rubber"] := "CARPET INSP"
lookup["Repair|Rubber"] := "RUBBER REP"
lookup["Refurb|Rubber"] := "RUBBER REFURB"
lookup["Demo|Rubber"] := "RUBBER DEMO"
lookup["Reinstall|Rubber"] := "RUBBER REINSTALL"
lookup["Prep|Rubber"] := "RUBBER PREP"
lookup["Adjusted Labor|Rubber"] := "RUBBER ADJ LABOR"
lookup["First time Install|VP Glue"] := "VP GLUE"
lookup["Reselect|VP Glue"] := "VP GLUE RESELECT"
lookup["Add on|VP Glue"] := "VP GLUE ADD ON"
lookup["Completion|VP Glue"] := "VP GLUE COMP"
lookup["Shortage|VP Glue"] := "VP GLUE SHORT"
lookup["Inspection|VP Glue"] := "VP GLUE INSP"
lookup["Repair|VP Glue"] := "VP GLUE REP"
lookup["Refurb|VP Glue"] := "VP GLUE REFURB"
lookup["Demo|VP Glue"] := "VP GLUE DEMO"
lookup["Reinstall|VP Glue"] := "VP GLUE REINSTALL"
lookup["Prep|VP Glue"] := "VP GLUE PREP"
lookup["Adjusted Labor|VP Glue"] := "VP GLUE ADJ LABOR"
lookup["First time Install|(Wood)"] := ""
lookup["Reselect|(Wood)"] := ""
lookup["Add on|(Wood)"] := ""
lookup["Completion|(Wood)"] := ""
lookup["Shortage|(Wood)"] := ""
lookup["Inspection|(Wood)"] := ""
lookup["Repair|(Wood)"] := ""
lookup["Refurb|(Wood)"] := ""
lookup["Demo|(Wood)"] := ""
lookup["Reinstall|(Wood)"] := ""
lookup["Prep|(Wood)"] := ""
lookup["Adjusted Labor|(Wood)"] := ""
lookup["First time Install|Gluedown wood"] := "GLUE WOOD"
lookup["Reselect|Gluedown wood"] := "GLUE WOOD RESELECT"
lookup["Add on|Gluedown wood"] := "GLUE WOOD ADD ON"
lookup["Completion|Gluedown wood"] := "GLUE WOOD COMP"
lookup["Shortage|Gluedown wood"] := "GLUE WOOD SHORT"
lookup["Inspection|Gluedown wood"] := "GLUE WOOD INSP"
lookup["Repair|Gluedown wood"] := "GLUE WOOD REP"
lookup["Refurb|Gluedown wood"] := "GLUE WOOD REFURB"
lookup["Demo|Gluedown wood"] := "GLUE WOOD DEMO"
lookup["Reinstall|Gluedown wood"] := "GLUE WOOD REINSTALL"
lookup["Prep|Gluedown wood"] := "GLUE WOOD PREP"
lookup["Adjusted Labor|Gluedown wood"] := "GLUE WOOD ADJ LABOR"
lookup["First time Install|(Carpet)"] := ""
lookup["Reselect|(Carpet)"] := ""
lookup["Add on|(Carpet)"] := ""
lookup["Completion|(Carpet)"] := ""
lookup["Shortage|(Carpet)"] := ""
lookup["Inspection|(Carpet)"] := ""
lookup["Repair|(Carpet)"] := ""
lookup["Refurb|(Carpet)"] := ""
lookup["Demo|(Carpet)"] := ""
lookup["Reinstall|(Carpet)"] := ""
lookup["Prep|(Carpet)"] := ""
lookup["Adjusted Labor|(Carpet)"] := ""
lookup["First time Install|Carpet"] := "CARPET"
lookup["Reselect|Carpet"] := "CARPET RESELECT"
lookup["Add on|Carpet"] := "CARPET ADD ON"
lookup["Completion|Carpet"] := "CARPET COMP"
lookup["Shortage|Carpet"] := "CARPET SHORT"
lookup["Inspection|Carpet"] := "CARPET INSP"
lookup["Repair|Carpet"] := "CARPET REP"
lookup["Refurb|Carpet"] := "CARPET REFURB"
lookup["Demo|Carpet"] := "CARPET DEMO"
lookup["Reinstall|Carpet"] := "CARPET REINSTALL"
lookup["Prep|Carpet"] := "CARPET PREP"
lookup["Adjusted Labor|Carpet"] := "CARPET ADJ LABOR"
lookup["First time Install|Carpet Square"] := "CARPET SQUARE"
lookup["Reselect|Carpet Square"] := "CARPET SQUARE RESELECT"
lookup["Add on|Carpet Square"] := "CARPET SQUARE ADD ON"
lookup["Completion|Carpet Square"] := "CARPET SQUARE COMP"
lookup["Shortage|Carpet Square"] := "CARPET SQUARE SHORT"
lookup["Inspection|Carpet Square"] := "CARPET INSP"
lookup["Repair|Carpet Square"] := "CARPET SQUARE REP"
lookup["Refurb|Carpet Square"] := "CARPET SQUARE REFURB"
lookup["Demo|Carpet Square"] := "CARPET SQUARE DEMO"
lookup["Reinstall|Carpet Square"] := "CARPET SQUARE REINSTALL"
lookup["Prep|Carpet Square"] := "CARPET SQUARE PREP"
lookup["Adjusted Labor|Carpet Square"] := "CARPET SQUARE ADJ LABOR"
lookup["First time Install|(Misc)"] := ""
lookup["Reselect|(Misc)"] := ""
lookup["Add on|(Misc)"] := ""
lookup["Completion|(Misc)"] := ""
lookup["Shortage|(Misc)"] := ""
lookup["Inspection|(Misc)"] := ""
lookup["Repair|(Misc)"] := ""
lookup["Refurb|(Misc)"] := ""
lookup["Demo|(Misc)"] := ""
lookup["Reinstall|(Misc)"] := ""
lookup["Prep|(Misc)"] := ""
lookup["Adjusted Labor|(Misc)"] := ""
lookup["First time Install|House file"] := "HOUSE FILE"
lookup["Reselect|House file"] := ""
lookup["Add on|House file"] := ""
lookup["Completion|House file"] := ""
lookup["Shortage|House file"] := ""
lookup["Inspection|House file"] := ""
lookup["Repair|House file"] := ""
lookup["Refurb|House file"] := ""
lookup["Demo|House file"] := ""
lookup["Reinstall|House file"] := ""
lookup["Prep|House file"] := ""
lookup["Adjusted Labor|House file"] := ""
lookup["First time Install|Slab"] := ""
lookup["Reselect|Slab"] := ""
lookup["Add on|Slab"] := ""
lookup["Completion|Slab"] := ""
lookup["Shortage|Slab"] := ""
lookup["Inspection|Slab"] := "SLAB INSP"
lookup["Repair|Slab"] := ""
lookup["Refurb|Slab"] := ""
lookup["Demo|Slab"] := ""
lookup["Reinstall|Slab"] := ""
lookup["Prep|Slab"] := ""
lookup["Adjusted Labor|Slab"] := ""
lookup["First time Install|Multiple"] := ""
lookup["Reselect|Multiple"] := ""
lookup["Add on|Multiple"] := ""
lookup["Completion|Multiple"] := ""
lookup["Shortage|Multiple"] := ""
lookup["Inspection|Multiple"] := "MULTI INSP"
lookup["Repair|Multiple"] := ""
lookup["Refurb|Multiple"] := ""
lookup["Demo|Multiple"] := ""
lookup["Reinstall|Multiple"] := ""
lookup["Prep|Multiple"] := ""
lookup["Adjusted Labor|Multiple"] := ""
lookup["First time Install|(Tile)"] := ""
lookup["Reselect|(Tile)"] := ""
lookup["Add on|(Tile)"] := ""
lookup["Completion|(Tile)"] := ""
lookup["Shortage|(Tile)"] := ""
lookup["Inspection|(Tile)"] := ""
lookup["Repair|(Tile)"] := ""
lookup["Refurb|(Tile)"] := ""
lookup["Demo|(Tile)"] := ""
lookup["Reinstall|(Tile)"] := ""
lookup["Prep|(Tile)"] := ""
lookup["Adjusted Labor|(Tile)"] := ""
lookup["First time Install|Multiple"] := ""
lookup["Reselect|Multiple"] := ""
lookup["Add on|Multiple"] := ""
lookup["Completion|Multiple"] := ""
lookup["Shortage|Multiple"] := ""
lookup["Inspection|Multiple"] := ""
lookup["Repair|Multiple"] := "TILE MULTI REP"
lookup["Refurb|Multiple"] := "TILE MULTI REFURB"
lookup["Demo|Multiple"] := "TILE MULTI DEMO"
lookup["Reinstall|Multiple"] := "TILE MULTI REINSTALL"
lookup["Prep|Multiple"] := ""
lookup["Adjusted Labor|Multiple"] := ""
lookup["First time Install|Tile Floor"] := "TILE FLOOR"
lookup["Reselect|Tile Floor"] := "TILE FLOOR RESELECT"
lookup["Add on|Tile Floor"] := "TILE FLOOR ADD ON"
lookup["Completion|Tile Floor"] := "TILE FLOOR COMP"
lookup["Shortage|Tile Floor"] := "TILE FLOOR SHORT"
lookup["Inspection|Tile Floor"] := "TILE INSP"
lookup["Repair|Tile Floor"] := "TILE FLOOR REP"
lookup["Refurb|Tile Floor"] := "TILE FLOOR REFURB"
lookup["Demo|Tile Floor"] := "TILE FLOOR DEMO"
lookup["Reinstall|Tile Floor"] := "TILE FLOOR REINSTALL"
lookup["Prep|Tile Floor"] := "TILE FLOOR PREP"
lookup["Adjusted Labor|Tile Floor"] := "TILE FLOOR ADJ LABOR"
lookup["First time Install|Tile Wall"] := "TILE WALL"
lookup["Reselect|Tile Wall"] := "TILE WALL RESELECT"
lookup["Add on|Tile Wall"] := "TILE WALL ADD ON"
lookup["Completion|Tile Wall"] := "TILE WALL COMP"
lookup["Shortage|Tile Wall"] := "TILE WALL SHORT"
lookup["Inspection|Tile Wall"] := "TILE INSP"
lookup["Repair|Tile Wall"] := "TILE WALL REP"
lookup["Refurb|Tile Wall"] := "TILE WALL REFURB"
lookup["Demo|Tile Wall"] := "TILE WALL DEMO"
lookup["Reinstall|Tile Wall"] := "TILE WALL REINSTALL"
lookup["Prep|Tile Wall"] := "TILE WALL PREP"
lookup["Adjusted Labor|Tile Wall"] := "TILE WALL ADJ LABOR"
lookup["First time Install|Tile Fireplace"] := "TILE FIREPLACE"
lookup["Reselect|Tile Fireplace"] := "TILE FIREPLACE RESELECT"
lookup["Add on|Tile Fireplace"] := "TILE FIREPLACE ADD ON"
lookup["Completion|Tile Fireplace"] := "TILE FIREPLACE COMP"
lookup["Shortage|Tile Fireplace"] := "TILE FIREPLACE SHORT"
lookup["Inspection|Tile Fireplace"] := "TILE INSP"
lookup["Repair|Tile Fireplace"] := "TILE FIREPLACE REP"
lookup["Refurb|Tile Fireplace"] := "TILE FIREPLACE REFURB"
lookup["Demo|Tile Fireplace"] := "TILE FIREPLACE DEMO"
lookup["Reinstall|Tile Fireplace"] := "TILE FIREPLACE REINSTALL"
lookup["Prep|Tile Fireplace"] := "TILE FIREPLACE PREP"
lookup["Adjusted Labor|Tile Fireplace"] := "TILE FIREPLACE ADJ LABOR"
lookup["First time Install|Tile Backsplash"] := "TILE BACKSPLASH"
lookup["Reselect|Tile Backsplash"] := "TILE BACKSPLASH RESELECT"
lookup["Add on|Tile Backsplash"] := "TILE BACKSPLASH ADD ON"
lookup["Completion|Tile Backsplash"] := "TILE BACKSPLASH COMP"
lookup["Shortage|Tile Backsplash"] := "TILE BACKSPLASH SHORT"
lookup["Inspection|Tile Backsplash"] := "TILE INSP"
lookup["Repair|Tile Backsplash"] := "TILE BACKSPLASH REP"
lookup["Refurb|Tile Backsplash"] := "TILE BACKSPLASH REFURB"
lookup["Demo|Tile Backsplash"] := "TILE BACKSPLASH DEMO"
lookup["Reinstall|Tile Backsplash"] := "TILE BACKSPLASH REINSTALL"
lookup["Prep|Tile Backsplash"] := "TILE BACKSPLASH PREP"
lookup["Adjusted Labor|Tile Backsplash"] := "TILE BACKSPLASH ADJ LABOR"
lookup["First time Install|Tile Balcony"] := "TILE BALCONY"
lookup["Reselect|Tile Balcony"] := "TILE BALCONY RESELECT"
lookup["Add on|Tile Balcony"] := "TILE BALCONY ADD ON"
lookup["Completion|Tile Balcony"] := "TILE BALCONY COMP"
lookup["Shortage|Tile Balcony"] := "TILE BALCONY SHORT"
lookup["Inspection|Tile Balcony"] := "TILE INSP"
lookup["Repair|Tile Balcony"] := "TILE BALCONY REP"
lookup["Refurb|Tile Balcony"] := "TILE BALCONY REFURB"
lookup["Demo|Tile Balcony"] := "TILE BALCONY DEMO"
lookup["Reinstall|Tile Balcony"] := "TILE BALCONY REINSTALL"
lookup["Prep|Tile Balcony"] := "TILE BALCONY PREP"
lookup["Adjusted Labor|Tile Balcony"] := "TILE BALCONY ADJ LABOR"
lookup["First time Install|Tile Patio"] := "TILE PATIO"
lookup["Reselect|Tile Patio"] := "TILE PATIO RESELECT"
lookup["Add on|Tile Patio"] := "TILE PATIO ADD ON"
lookup["Completion|Tile Patio"] := "TILE PATIO COMP"
lookup["Shortage|Tile Patio"] := "TILE PATIO SHORT"
lookup["Inspection|Tile Patio"] := "TILE INSP"
lookup["Repair|Tile Patio"] := "TILE PATIO REP"
lookup["Refurb|Tile Patio"] := "TILE PATIO REFURB"
lookup["Demo|Tile Patio"] := "TILE PATIO DEMO"
lookup["Reinstall|Tile Patio"] := "TILE PATIO REINSTALL"
lookup["Prep|Tile Patio"] := "TILE PATIO PREP"
lookup["Adjusted Labor|Tile Patio"] := "TILE PATIO ADJ LABOR"
lookup["First time Install|Tile Pavers"] := "TILE PAVERS"
lookup["Reselect|Tile Pavers"] := "TILE PAVERS RESELECT"
lookup["Add on|Tile Pavers"] := "TILE PAVERS ADD ON"
lookup["Completion|Tile Pavers"] := "TILE PAVERS COMP"
lookup["Shortage|Tile Pavers"] := "TILE PAVERS SHORT"
lookup["Inspection|Tile Pavers"] := "TILE INSP"
lookup["Repair|Tile Pavers"] := "TILE PAVERS REP"
lookup["Refurb|Tile Pavers"] := "TILE PAVERS REFURB"
lookup["Demo|Tile Pavers"] := "TILE PAVERS DEMO"
lookup["Reinstall|Tile Pavers"] := "TILE PAVERS REINSTALL"
lookup["Prep|Tile Pavers"] := "TILE PAVERS PREP"
lookup["Adjusted Labor|Tile Pavers"] := "TILE PAVERS ADJ LABOR"
lookup["First time Install|Grout"] := ""
lookup["Reselect|Grout"] := "GROUT RESELECT"
lookup["Add on|Grout"] := ""
lookup["Completion|Grout"] := ""
lookup["Shortage|Grout"] := "GROUT SHORT"
lookup["Inspection|Grout"] := "GROUT INSP"
lookup["Repair|Grout"] := "GROUT REP"
lookup["Refurb|Grout"] := ""
lookup["Demo|Grout"] := ""
lookup["Reinstall|Grout"] := ""
lookup["Prep|Grout"] := ""
lookup["Adjusted Labor|Grout"] := ""
lookup["First time Install|Caulking"] := "CAULKING"
lookup["Reselect|Caulking"] := ""
lookup["Add on|Caulking"] := ""
lookup["Completion|Caulking"] := "CAULKING COMP"
lookup["Shortage|Caulking"] := ""
lookup["Inspection|Caulking"] := ""
lookup["Repair|Caulking"] := "CAULKING REP"
lookup["Refurb|Caulking"] := "CAULKING REFURB"
lookup["Demo|Caulking"] := ""
lookup["Reinstall|Caulking"] := "CAULKING REFURB"
lookup["Prep|Caulking"] := ""
lookup["Adjusted Labor|Caulking"] := ""
lookup["First time Install|(Separated)"] := ""
lookup["Reselect|(Separated)"] := ""
lookup["Add on|(Separated)"] := ""
lookup["Completion|(Separated)"] := ""
lookup["Shortage|(Separated)"] := ""
lookup["Inspection|(Separated)"] := ""
lookup["Repair|(Separated)"] := ""
lookup["Refurb|(Separated)"] := ""
lookup["Demo|(Separated)"] := ""
lookup["Reinstall|(Separated)"] := ""
lookup["Prep|(Separated)"] := ""
lookup["Adjusted Labor|(Separated)"] := ""
lookup["First time Install|Waterproof"] := "WATERPROOF"
lookup["Reselect|Waterproof"] := ""
lookup["Add on|Waterproof"] := ""
lookup["Completion|Waterproof"] := "WATERPROOF COMP"
lookup["Shortage|Waterproof"] := ""
lookup["Inspection|Waterproof"] := ""
lookup["Repair|Waterproof"] := "WATERPROOF REP"
lookup["Refurb|Waterproof"] := "WATERPROOF REFURB"
lookup["Demo|Waterproof"] := ""
lookup["Reinstall|Waterproof"] := "WATERPROOF REFURB"
lookup["Prep|Waterproof"] := ""
lookup["Adjusted Labor|Waterproof"] := ""
lookup["First time Install|Tile Wall Master"] := "TILE WALL MST"
lookup["Reselect|Tile Wall Master"] := "TILE WALL MST RESELECT"
lookup["Add on|Tile Wall Master"] := "TILE WALL MST ADD ON"
lookup["Completion|Tile Wall Master"] := "TILE WALL COMP"
lookup["Shortage|Tile Wall Master"] := "TILE WALL SHORT"
lookup["Inspection|Tile Wall Master"] := "TILE INSP"
lookup["Repair|Tile Wall Master"] := "TILE WALL REP"
lookup["Refurb|Tile Wall Master"] := "TILE WALL REFURB"
lookup["Demo|Tile Wall Master"] := "TILE WALL DEMO"
lookup["Reinstall|Tile Wall Master"] := "TILE WALL REINSTALL"
lookup["Prep|Tile Wall Master"] := "TILE WALL PREP"
lookup["Adjusted Labor|Tile Wall Master"] := "TILE WALL ADJ LABOR"
lookup["First time Install|Tile Wall Secondary"] := "TILE WALL SEC"
lookup["Reselect|Tile Wall Secondary"] := "TILE WALL SEC RESELECT"
lookup["Add on|Tile Wall Secondary"] := "TILE WALL SEC ADD ON"
lookup["Completion|Tile Wall Secondary"] := "TILE WALL COMP"
lookup["Shortage|Tile Wall Secondary"] := "TILE WALL SHORT"
lookup["Inspection|Tile Wall Secondary"] := "TILE INSP"
lookup["Repair|Tile Wall Secondary"] := "TILE WALL REP"
lookup["Refurb|Tile Wall Secondary"] := "TILE WALL REFURB"
lookup["Demo|Tile Wall Secondary"] := "TILE WALL DEMO"
lookup["Reinstall|Tile Wall Secondary"] := "TILE WALL REINSTALL"
lookup["Prep|Tile Wall Secondary"] := "TILE WALL PREP"
lookup["Adjusted Labor|Tile Wall Secondary"] := "TILE WALL ADJ LABOR"
lookup["First time Install|Tile Wainscot"] := "TILE WAINSCOT"
lookup["Reselect|Tile Wainscot"] := "TILE WAINSCOT RESELECT"
lookup["Add on|Tile Wainscot"] := "TILE WAINSCOT ADD ON"
lookup["Completion|Tile Wainscot"] := "TILE WALL COMP"
lookup["Shortage|Tile Wainscot"] := "TILE WALL SHORT"
lookup["Inspection|Tile Wainscot"] := "TILE INSP"
lookup["Repair|Tile Wainscot"] := "TILE WALL REP"
lookup["Refurb|Tile Wainscot"] := "TILE WALL REFURB"
lookup["Demo|Tile Wainscot"] := "TILE WALL DEMO"
lookup["Reinstall|Tile Wainscot"] := "TILE WALL REINSTALL"
lookup["Prep|Tile Wainscot"] := "TILE WALL PREP"
lookup["Adjusted Labor|Tile Wainscot"] := "TILE WALL ADJ LABOR"
lookup["First time Install|Tile Tub Area"] := "TILE TUB AREA"
lookup["Reselect|Tile Tub Area"] := "TILE TUB AREA RESELECT"
lookup["Add on|Tile Tub Area"] := "TILE TUB AREA ADD ON"
lookup["Completion|Tile Tub Area"] := "TILE TUB AREA COMP"
lookup["Shortage|Tile Tub Area"] := "TILE TUB AREA SHORT"
lookup["Inspection|Tile Tub Area"] := "TILE INSP"
lookup["Repair|Tile Tub Area"] := "TILE TUB AREA REP"
lookup["Refurb|Tile Tub Area"] := "TILE TUB AREA REFURB"
lookup["Demo|Tile Tub Area"] := "TILE TUB AREA DEMO"
lookup["Reinstall|Tile Tub Area"] := "TILE TUB AREA REINSTALL"
lookup["Prep|Tile Tub Area"] := "TILE WALL PREP"
lookup["Adjusted Labor|Tile Tub Area"] := "TILE WALL ADJ LABOR"
lookup["First time Install|Tile Seat"] := "TILE SEAT"
lookup["Reselect|Tile Seat"] := "TILE SEAT RESELECT"
lookup["Add on|Tile Seat"] := "TILE SEAT ADD ON"
lookup["Completion|Tile Seat"] := "TILE SEAT COMP"
lookup["Shortage|Tile Seat"] := "TILE SEAT SHORT"
lookup["Inspection|Tile Seat"] := "TILE INSP"
lookup["Repair|Tile Seat"] := "TILE SEAT REP"
lookup["Refurb|Tile Seat"] := "TILE SEAT REFURB"
lookup["Demo|Tile Seat"] := "TILE SEAT DEMO"
lookup["Reinstall|Tile Seat"] := "TILE SEAT REINSTALL"
lookup["Prep|Tile Seat"] := "TILE WALL PREP"
lookup["Adjusted Labor|Tile Seat"] := "TILE WALL ADJ LABOR"
lookup["First time Install|Tile Mudpan"] := "TILE MUDPAN"
lookup["Reselect|Tile Mudpan"] := "TILE MUDPAN RESELECT"
lookup["Add on|Tile Mudpan"] := "TILE MUDPAN ADD ON"
lookup["Completion|Tile Mudpan"] := "TILE MUDPAN COMP"
lookup["Shortage|Tile Mudpan"] := "TILE MUDPAN SHORT"
lookup["Inspection|Tile Mudpan"] := "TILE INSP"
lookup["Repair|Tile Mudpan"] := "TILE MUDPAN REP"
lookup["Refurb|Tile Mudpan"] := "TILE MUDPAN REFURB"
lookup["Demo|Tile Mudpan"] := "TILE MUDPAN DEMO"
lookup["Reinstall|Tile Mudpan"] := "TILE MUDPAN REINSTALL"
lookup["Prep|Tile Mudpan"] := "TILE WALL PREP"
lookup["Adjusted Labor|Tile Mudpan"] := "TILE WALL ADJ LABOR"
lookup["First time Install|Tile Dog Wash"] := "TILE DOG WASH"
lookup["Reselect|Tile Dog Wash"] := "TILE DOG WASH RESELECT"
lookup["Add on|Tile Dog Wash"] := "TILE DOG WASH ADD ON"
lookup["Completion|Tile Dog Wash"] := "TILE DOG WASH COMP"
lookup["Shortage|Tile Dog Wash"] := "TILE DOG WASH SHORT"
lookup["Inspection|Tile Dog Wash"] := "TILE INSP"
lookup["Repair|Tile Dog Wash"] := "TILE DOG WASH REP"
lookup["Refurb|Tile Dog Wash"] := "TILE DOG WASH REFURB"
lookup["Demo|Tile Dog Wash"] := "TILE DOG WASH DEMO"
lookup["Reinstall|Tile Dog Wash"] := "TILE DOG WASH REINSTALL"
lookup["Prep|Tile Dog Wash"] := "TILE WALL PREP"
lookup["Adjusted Labor|Tile Dog Wash"] := "TILE WALL ADJ LABOR"

overrideMap := Map()
overrideMap["ACCLIMATION|TILE"] := ""
overrideMap["ACID WASH|TILE"] := "ACID WASH"
overrideMap["GROUT STAIN|TILE"] := "GROUT STAIN"
overrideMap["MATERIAL DELIVERY|TILE"] := "MATERIAL DELIVERY"
overrideMap["MATERIAL PICK UP|TILE"] := "MATERIAL PICK UP"
overrideMap["MATERIAL NOT RETURNED|TILE"] := "MATERIAL NOT RETURNED"
overrideMap["DRAW ONLY|TILE"] := "TILE DRAW"
overrideMap["MOVERS|TILE"] := "MOVERS"
overrideMap["TRIP CHARGE|TILE"] := "TILE TRIP"
overrideMap["REIMBURSEMENT|TILE"] := "REIMBURSEMENT"
overrideMap["HO SETTLEMENT|TILE"] := "HO SETTLEMENT"
overrideMap["ACCLIMATION|VINYL"] := "ACCLIMATION"
overrideMap["ACID WASH|VINYL"] := ""
overrideMap["GROUT STAIN|VINYL"] := ""
overrideMap["MATERIAL DELIVERY|VINYL"] := "MATERIAL DELIVERY"
overrideMap["MATERIAL PICK UP|VINYL"] := "MATERIAL PICK UP"
overrideMap["MATERIAL NOT RETURNED|VINYL"] := "MATERIAL NOT RETURNED"
overrideMap["DRAW ONLY|VINYL"] := "VINYL DRAW"
overrideMap["MOVERS|VINYL"] := "MOVERS"
overrideMap["TRIP CHARGE|VINYL"] := "VINYL TRIP"
overrideMap["REIMBURSEMENT|VINYL"] := "REIMBURSEMENT"
overrideMap["HO SETTLEMENT|VINYL"] := "HO SETTLEMENT"
overrideMap["ACCLIMATION|WOOD"] := "ACCLIMATION"
overrideMap["ACID WASH|WOOD"] := ""
overrideMap["GROUT STAIN|WOOD"] := ""
overrideMap["MATERIAL DELIVERY|WOOD"] := "MATERIAL DELIVERY"
overrideMap["MATERIAL PICK UP|WOOD"] := "MATERIAL PICK UP"
overrideMap["MATERIAL NOT RETURNED|WOOD"] := "MATERIAL NOT RETURNED"
overrideMap["DRAW ONLY|WOOD"] := "WOOD DRAW"
overrideMap["MOVERS|WOOD"] := "MOVERS"
overrideMap["TRIP CHARGE|WOOD"] := "WOOD TRIP"
overrideMap["REIMBURSEMENT|WOOD"] := "REIMBURSEMENT"
overrideMap["HO SETTLEMENT|WOOD"] := "HO SETTLEMENT"
overrideMap["ACCLIMATION|CARPET"] := ""
overrideMap["ACID WASH|CARPET"] := ""
overrideMap["GROUT STAIN|CARPET"] := ""
overrideMap["MATERIAL DELIVERY|CARPET"] := "MATERIAL DELIVERY"
overrideMap["MATERIAL PICK UP|CARPET"] := "MATERIAL PICK UP"
overrideMap["MATERIAL NOT RETURNED|CARPET"] := "MATERIAL NOT RETURNED"
overrideMap["DRAW ONLY|CARPET"] := "CARPET DRAW"
overrideMap["MOVERS|CARPET"] := "MOVERS"
overrideMap["TRIP CHARGE|CARPET"] := "CARPET TRIP"
overrideMap["REIMBURSEMENT|CARPET"] := "REIMBURSEMENT"
overrideMap["HO SETTLEMENT|CARPET"] := "HO SETTLEMENT"

UpdateResult() {

    selectedPrefix := cbPrefix.Text

    prefix := prefixMap.Has(selectedPrefix) ? prefixMap[selectedPrefix] : ""

    override := cbOverride.Value

    if override {

        category := cbCategory.Text

        material := cbMaterial.Text

        key := category "|" material

        result := overrideMap.Has(key) ? overrideMap[key] : ""

    } else {

        selectedStage := cbStage.Text
        selectedCommodity := cbCommodity.Text
        key := selectedStage "|" selectedCommodity
        result := lookup.Has(key) ? lookup[key] : ""

    }

    tbResult.Value := Trim(StrUpper(prefix " " result))

}
