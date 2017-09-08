__author__ = 'David Oreper'
class PeopleDataKeys:
    def __init__(self):
        pass

    MANAGER = "Manager"
    NAME = "HR Name"
    NICK_NAME = "Nickname"
    LEVEL = "Level"
    FUNCTION = "Function"
    PROJECT = "Project"
    FEATURE_TEAM = "Feature Team"
    TYPE = "Type"
    REQ = "Requisition Number"
    CONSULTANT = "Consultant"
    CONTRACTOR = "Contractor"
    EXPAT_TYPE = "Expat"
    VENDOR_TYPE = "Vendor"
    INTERN_TYPE = "Intern"
    LOCATION = "Location"
    START_DATE = "Start Date"

    CROSS_FUNCTIONS = ["admin", "admin operations", "devops","inf", "infrastructure", "cross functional", "customer success", "technology",]
    CROSS_FUNCT_TEAM = "cross"
    FLOORS = {}
    TEAM_MODEL = {}
    PRODUCT_SORT_ORDER = []
    FLOOR_SORT_ORDER = []

class PeopleDataKeysBellevue(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)

    # CROSS_FUNCTIONS = ["technology", "admin", "inf", "infrastructure", "cross functional"]

class PeopleDataKeysSantaClara(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)

    LEVEL = "Title"
    TEAM_MODEL = {
    "UCP" : "1 Tracks @ (1 PO, 1 TA,  4 Dev, 1 QA, 2 Char, 2 Auto)",
    "HID" : "2 Tracks @ (1 PO, 5 Dev, 2 QA, 2 Auto, 1 UX)",
    "HVS" : "Q1:20; Q2:25; Q3:27; Q4:32 -- 1 Tracks @ (1 PO, 5 Dev, 1 QA, 1 Auto)",
    "Evidence Management" : "1 Tracks @ (1 PO, 4 Dev, 1 QA, 1 Auto)",
    "HCmD" : "1 Tracks @ (1 Head Coach, 2 PO, 2 Dev, 1 QA, 1 UX)",

    }


class PeopleDataKeysSIBU(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)
    REQ = "Requisition"
    LEVEL = "Title"
    TEAM_MODEL = {
            "HVS" : "[Forecast: Q1:20; Q2:26; Q3:29; Q4:34] -- 1 Tracks @ (1 PO, 5 Dev, 1 QA, 1 Auto)",
            "HVS EM" : "2 Tracks @ (1 PO, 4 Dev, 1 QA, 1 Char, 1 Auto)",
            "Lumada - System" : "[Forecast: Q1:7; Q2:6; Q3:43; Q4:110]",
            "Lumada - Studio" : "[Forecast: Q1:7; Q2:10; Q3:43; Q4:110]",
            "City Data Exchange" : "[Forecast: Q1:6; Q2:19; Q3:6; Q4:6]",
            "Predictive Maintenance" : "[Forecast: Q1:5; Q2:17; Q3:22; Q4:27]",
            "Optimized Factory" : "[Forecast: Q1:1; Q2:6; Q3:13; Q4:15]",
        }

    PRODUCT_SORT_ORDER = ["hvs", "hvs em", "vmp", "hvp", "smart city technology", "technology", "tactical integration",
                          "tactical integrations", "lumada - system", "sc iiot", "set", "bel iiot", "lumada platform", "pdm", "predictive maintenance",
                          "lumada - studio", "lumada - microservices", "optimized factory", "opf", "city data exchange",
                          "cde", "denver", "lumada - ai", "lumada - analytics", "lumada - di", "lumada - hci", "hci", "lumada - machine intelligence", "lumada", "cross", "lumada cross", "global"]

class PeopleDataKeysWaltham(PeopleDataKeysSIBU):
    def __init__(self):
        PeopleDataKeys.__init__(self)
    FUNCTION = "Function"
    # CROSS_FUNCTIONS = ["Technology", "DevOps", "Admin", "Sustaining" ]
    FLOORS = {
        "- Mobility": [
            "Anderson, Vic",
            "Kostadinov, Alex",
            "Lin, Wayzen",
            "Manjanatha, Sowmya",
            "Maruca, Fran",
            "Pfahl, Matt",
            "Van Thong, Adrien",
        ],


        "- Content": [
            "Boba, Andrew",
            "Bronner, Mark",
            "Burnham, John",
            "Chestna, Wayne",
            "Lee, Jonathan",
            "Shea, Kevin",
        ],

        "- Aspen": [
            "Hartford, Joe",
            "Liang, Candy",
        ],

        "- HPP": [
            "Wesley, Joe",
            "Moore, Jim",
        ],
        "- HDID": [
            "Agashe, Sujata",
            "Caswell, Paul",
            "Chappell, Simon",
            "Gothoskar, Chandrashekhar",
            "Helliker, Fabrice",
            "Mason, Bill",
            "Melville, Andrew",
            "Pendlebury, Ian",
            "Pfaff, Florian",
            "Sinkar, Milind",
        ],
    }

    TEAM_MODEL = {
        "Aspen" : "4 Tracks @ (1 PO, 3 Dev, 1 QA, 1 Char, 1 Auto)",
        "HCP-Rhino" : "4 Tracks @ (1 PO, 4 Dev, 2 QA, 1 Char, 2 Auto)",
        "HCP-India" : "1 Track @ (1 PO, 3 Dev, 1 QA, 1 Auto)",
#        "HCP (Rhino)" : "1 Track @ (1 PO, 4 Dev, 2 QA, 2 Char, 2 Auto)",
        "HCP-AW" : "4 Tracks @ (1 PO, 4 Dev, 2 QA, 1 Char, 2 Auto)",
        }

    # names should be lower case here
    PRODUCT_SORT_ORDER = ["aspen", "ensemble", "hcp-rhino", "hcp-india", "hcp-aw", "aw-japan","hpp", "hpp-india", "hdid-uk", "hdid-waltham", "hdid-germany", "hdid-pune", "future funding"]
    FLOOR_SORT_ORDER = ["- ensemble", "- content", "- mobility", "- hpp" ]


class PeopleDataKeysSIBU(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)
    REQ = "Requisition"
    LEVEL = "Title"
    TEAM_MODEL = {
            "HVS" : "[Forecast: Q1:20; Q2:26; Q3:29; Q4:34] -- 1 Tracks @ (1 PO, 5 Dev, 1 QA, 1 Auto)",
            "HVS EM" : "2 Tracks @ (1 PO, 4 Dev, 1 QA, 1 Char, 1 Auto)",
            "Lumada - System" : "[Forecast: Q1:7; Q2:6; Q3:43; Q4:110]",
            "Lumada - Studio" : "[Forecast: Q1:7; Q2:10; Q3:43; Q4:110]",
            "City Data Exchange" : "[Forecast: Q1:6; Q2:19; Q3:6; Q4:6]",
            "Predictive Maintenance" : "[Forecast: Q1:5; Q2:17; Q3:22; Q4:27]",
            "Optimized Factory" : "[Forecast: Q1:1; Q2:6; Q3:13; Q4:15]",
        }

    PRODUCT_SORT_ORDER = ["hvs", "hvs em", "vmp", "hvp", "smart city technology", "technology", "tactical integration",
                          "tactical integrations",  "lumada - system", "sc iiot", "bel iiot", "lumada platform", "pdm", "predictive maintenance",
                          "lumada - studio", "lumada - microservices", "optimized factory", "opf", "city data exchange",
                          "cde", "denver", "lumada - ai", "lumada - analytics", "lumada - di", "lumada - hci", "hci", "lumada - machine intelligence", "lumada", "cross", "lumada cross", "global"]

class PeopleDataKeysHPP(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)
    FUNCTION = "Function"
    #CROSS_FUNCTIONS = ["Technology", "DevOps", "Admin", "Seal" ]
