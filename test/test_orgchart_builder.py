import pprint
from unittest import TestCase
from orgchart_parser import OrgParser, PeopleDataKeys
import orgchart_parser

__author__ = 'doreper'


class TestOrgChartBuilder(TestCase):
    def setUp(self):
        self.waltham = "./Waltham Staff.xlsm"
        self.bellevue = "./Bellevue Staff.xlsm"
        self.santaClara = "./SantaClara Staff.xlsm"
        self.sheetName = "PeopleData"

    def test_getWalthamManagers(self):
        dataSheet = OrgParser(self.waltham, self.sheetName)
        self.assertEqual({u'Anderson, Vic', u'Burnham, John', u'Rogers, Rich', u'Pinkney, Dave', u'Chestna, Wayne',
                          u'Isherwood, Ben', u'Pannese, Donald', u'Liang, Candy'}, dataSheet.getManagerSet())

    def test_getWalthamProducts(self):
        dataSheet = OrgParser(self.waltham, self.sheetName)
        self.assertEqual({u'Product Team', u'HCP', u'HDDS', u'DevOps', u'Free', u'Cross', u'HCP-AW', u'Tech',
                          u'HCP (Rhino)'}, dataSheet.getProductSet())

    def test_getWalthamFunction(self):
        dataSheet = OrgParser(self.waltham, self.sheetName)
        self.assertEqual({u'Functional Team', u'Lead', u'Characterization', u'Auto', u'Dev', u'QA', u'UI',
                          u'Sustaining', u'PO', u'Tech', u'DevOps', u'Free'}, dataSheet.getFunctionSet())

    def test_getWalthamAdminDict(self):
        dataSheet = OrgParser(self.waltham, self.sheetName)
        pprint.pprint(dataSheet.getAdminDict())
        expectedAdminDict = {'Anderson, Vic': ['Brandon Bateman<br>Senior, SW Engineer**',
                   'Beth Tirado<br>Specialist, SW Engineer',
                   'Hank Wilder<br>Specialist, SWQA Engineer'],
 'Burnham, John': ['Brendan Almonte<br>Software Engineer',
                   'Sameer Apte<br>Specialist, SW Engineer',
                   'Andrew Todd<br>Senior, SW Engineer',
                   'Ngale Truong<br>Senior, SW Engineer',
                   'Adrien Van Thong<br>Senior, SW Engineer'],
 'Chestna, Wayne': ['Emily Wilson<br>SWQA Engineer'],
 'Isherwood, Ben': ['Ilya Tatar<br>Specialist, SW Engineer'],
 'Liang, Candy': ['Logan Stuart<br>Associate, SW Engineer',
                  'Nitesh Taneja<br>Specialist, SW Engineer',
                  'Eric Wilson<br>Senior, SW Engineer'],
 'Pannese, Donald': ['Tom Baron<br>Specialist, SWQA Engineer'],
 'Pinkney, Dave': ['Walter Wohler<br>Software Engineer',
                   'Scott Yaninas<br>Software Engineer',
                   'Vitaly Zolotusky<br>Master, SW Engineer'],
 'Rogers, Rich': ['Vic Anderson<br>Senior Director, Engineering=='],
 '_Manager Not Set': ['TBH15<br>13240<br>Senior, SW Engineer',
                      'TBH8<br>13239<br>Senior, SW Engineer',
                      'TBH21<br>13237<br>Specialist, SW Engineer',
                      'TBH22<br>13238<br>Specialist, SW Engineer',
                      'TBH3<br>13171<br>Software Engineer',
                      'TBH5<br>13172<br>Senior, SW Engineer**',
                      'TBH6<br>Specialist, SW Engineer',
                      'TBH10<br>12829<br>Senior, SW Engineer',
                      'TBH1<br>12372<br>Software Engineer**',
                      'TBH2<br>11952<br>Associate, SW Engineer**',
                      'TBH23<br>12455<br>Specialist, SW Engineer',
                      'TBH7<br>Software Engineer**',
                      'TBH24<br>Software Engineer**',
                      'TBH12<br>12706<br>Software Engineer',
                      'TBH13<br>11940<br>Senior Manager, Engineering',
                      'TBH11<br>11948<br>Senior, SW Engineer',
                      'TBH14<br>11957<br>Senior, SW Engineer**',
                      'Zhiyang Tan<br>Associate, SW Engineer**',
                      'TBH16<br>13363<br>Software Engineer**',
                      'TBH20<br>13167<br>Senior, SWQA Engineer**',
                      'TBH9<br>12887<br>Software Engineer**',
                      'Name',
                      'Kyosuke Achiwa',
                      'Keita Hosoi',
                      'Masahiro Shimizu',
                      'Bryan Yergeau',
                      'Julie Brady',
                      'Tyler Wright']}
        self.maxDiff = None
        self.assertEqual(expectedAdminDict, dataSheet.getAdminDict())

    def test_getBellevueManagers(self):
        dataSheet = OrgParser(self.bellevue, self.sheetName)
        self.assertEqual({u'Gaurav Bora', u'Abhinaw Dixit', u'Jeb Garcia'}, dataSheet.getManagerSet())

    def test_getBellevueProducts(self):
        dataSheet = OrgParser(self.bellevue, self.sheetName)
        self.assertEqual({u'Shasta', u'UCP', u'Rainier'}, dataSheet.getProductSet())

    def test_getBellevueFunction(self):
        dataSheet = OrgParser(self.bellevue, self.sheetName)
        self.assertEqual({u'Development', u'Stress', u'SW Architecture', u'Product Owner', u'Automation',
                          u'Solutions and Sustaining', u'Technology'}, dataSheet.getFunctionSet())

    def test_getBellevueAdminDict(self):
        dataSheet = OrgParser(self.bellevue, self.sheetName)
        expectedAdminDict = {'Abhinaw Dixit': ['Calvin Lewis<br>Software Engineer**'],
                             'Gaurav Bora': ['Bill Zietzke<br>Master, SW Engineer',
                                             'Jeb Garcia<br>Manager, SWQA=='],
                             'Jeb Garcia': ['Sean Kim<br>Associate, SWQA Engineer**',
                                            'Ying Xue<br>Associate, SW Engineer**',
                                            'Tony She<br>Associate, SWQA Engineer**',
                                            'Devan Tatum<br>Associate, SWQA Engineer'],
                             '_Manager Not Set': ['Rich Rogers<br>Senior VP, Engineering',
                                                  'Ankit Arora<br>Intern**',
                                                  'Trent McFarlane<br>Intern**',
                                                  'TBH 1',
                                                  'TBH 2',
                                                  'TBH 3',
                                                  'TBH 4',
                                                  'TBH 5',
                                                  'TBH 6',
                                                  'TBH 7',
                                                  'TBH 8',
                                                  'TBH 9']}
        pprint.pprint(dataSheet.getAdminDict())
        self.maxDiff = None
        self.assertEqual(expectedAdminDict, dataSheet.getAdminDict())

    def test_getSCFuncDict(self):
        dataSheet = OrgParser(self.santaClara, self.sheetName)
        expectedStr = """Cross\r
	Inf\r
		Barry Van Hooser<br>Specialist, Solutions Architect\r
		Krishna Botlagunta\r
		Logan Hawkes<br>Specialist, SWQA Engineer\r
		Sarika Donakanti\r
Expat\r
	Dev\r
		Naoya Murao (HCmD)\r
	QA\r
		Susumu Tomita (Rainier)\r
HCmD\r
	PO\r
		Arno Grbac<br>Director, Engineering==\r
		Sanjay Sharma<br>Specialist, SW Engineer==\r
		Varun Sood<br>Senior, SW Engineer**\r
	Dev\r
		Farida Fatehi<br>Senior, SW Engineer\r
		Nischitha Puttaswamy<br>Software Engineer**\r
		Rajiv Rajput<br>Senior, SW Engineer**\r
		Sameer Vulchi<br>Senior Manager, Engineering==\r
		Sathish Raghunathan<br>Senior, SW Engineer\r
	QA\r
		Abhinay Wankhade<br>Manager, Quality Assurance==\r
		Malya Das**\r
		Manish Joshi<br>Senior, SWQA Engineer**\r
		Sandy Ersheid<br>Specialist, SWQA Engineer\r
		Sheela Shivayogi<br>Senior, SWQA Engineer\rpa60t@111
		Srinivas Abburi<br>Specialist, SWQA Engineer**\r
	Aut\r
		Chongrui Duan<br>Associate, SW Engineer**\r
		Denis Molchanenko<br>Specialist, SW Engineer**\r
		Dian Wang<br>Associate, SW Engineer**\r
		Kenneth Fung<br>Senior, SW Engineer\r
		Ramin Tawakuli<br>Associate, SW Engineer**\r
		Siarhei Tolkach<br>Senior, SW Engineer**\r
	UX\r
		Abdul Athaullah<br>Senior, SW Engineer\r
Rainier\r
	PO\r
		Ganesh Kaliamourthy<br>Senior Manager, Engineering==\r
		Meghana Janumpally<br>Senior, SW Engineer\r
	Dev\r
		Ankur Avlani<br>Senior, SW Engineer\r
		Nitin Wilson<br>Manager, Software Development==\r
		Prateek Demla<br>Software Engineer**\r
		Rishik Dhar<br>Senior, SW Engineer\r
		Scott Kawaguchi<br>Senior, SW Engineer\r
		Shraddha Herlekar<br>Software Engineer**\r
		Supriya Grandhi<br>Software Engineer\r
	QA\r
		Aparna Tengse<br>Senior, SWQA Engineer**\r
		Favad Khan<br>Senior, SWQA Engineer\r
		Jyoti Suryaji<br>Senior, SWQA Engineer\r
	Aut\r
		John Colarusso<br>Associate, SW Engineer**\r
		Sriram Venkataraman<br>Senior, SW Engineer\r
	UX\r
		Alena Starostina<br>Senior, Product Management Analyst\r
		Tiana Cavros<br>Associate, Product Management Analyst\r
UCP\r
	PO\r
		TBH<br>13217\r
		TBH<br>13218<br>Senior, SW Engineer\r
		TBH<br>13219<br>Senior, SW Engineer\r
		TBH<br>13220<br>Senior, SW Engineer\r
		TBH<br>13221<br>Senior, SW Engineer\r
		TBH<br>13222<br>Senior, SW Engineer\r
	PO\r
		TBH<br>13223<br>Senior, SW Engineer\r
		TBH<br>13224<br>Senior, SW Engineer\r
		TBH<br>13225<br>Senior, SW Engineer\r
		TBH<br>13226<br>Senior, SW Engineer\r
		TBH<br>13234<br>Senior, SW Engineer\r
		TBH<br>13235<br>Senior, SW Engineer\r
	PO\r
		TBH<br>13236<br>Senior, SW Engineer\r
		TBH<br>13237<br>Senior, SW Engineer\r
		TBH<br>13238<br>Senior, SW Engineer\r
		TBH<br>13239\r
	TA\r
		Andrew Nielsen<br>Director, Technology\r
	Dev\r
		Adelle Knight<br>Specialist, SW Engineer**\r
		Aruna Gummalla<br>Specialist, SW Engineer**\r
		Muralidhar Chapa<br>Specialist, SW Engineer\r
		Shaikh Wasiullah<br>Senior, SW Engineer\r
		Sumeet Mittal<br>Specialist, SW Engineer**\r
	QA\r
		Balaji Ramanuja<br>Senior, SWQA Engineer\r
		Jagriti Wadhwa<br>SWQA Engineer**\r
	Aut\r
		David Oreper<br>Manager, Engineering==\r
		John-Paul Victoria<br>Senior, SW Engineer\r
	Test\r
		TBH<br>13227<br>Senior, SW Engineer\r
		TBH<br>13228<br>Senior, SW Engineer\r
		TBH<br>13229<br>Senior, SW Engineer\r
		TBH<br>13230<br>Senior, SW Engineer\r
		TBH<br>13231<br>Senior, SW Engineer\r
		TBH<br>13232<br>Senior, SW Engineer\r
		TBH<br>13233<br>Senior, SW Engineer\r
"""
        print(dataSheet._getFormattedFuncStr())

        self.assertEqual(expectedStr, dataSheet._getFormattedFuncStr())

    def test_filename_func(self):
        orgchart_parser.main("vue staff.xl -f".split(" "))