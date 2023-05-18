"""
Name:Francis Santos
Start date: 15/05/2023
Finish date: /05/2023
Title: Unit Test with normal code and exception handling
"""
import unittest

from WhiteBoxControlCode import *

class UnitTestTimetableApp(unittest.TestCase):
    
    def testsetupcounter_1(self):
        counter = 0
        file_loc = countertest(counter)
        self.assertEqual(file_loc,'C:\\Users\\FrancisS\\Downloads\\AnonymisedSPlusData.xlsx')
        
    def testsetupcounter_2(self):
        counter = 1
        file_loc = countertest(counter)
        self.assertEqual(file_loc,'C:\\Users\\FrancisS\\Downloads\\Copyofanonymised_names1.xlsx')
        
    def testsetupdefault(self):
        file_loc = testdefault()
        self.assertEqual(file_loc, r'C:\Users\FrancisS\Downloads\Copyofanonymised_names1.xlsx')

    def testprocess1ex1(self):
        dataframe = 'Weeks 4-16'
        result = tesprocessF1(dataframe)
        self.assertEqual(result, True)

    def testprocess1ex2(self):
        dataframe = 'Term 1'
        result = tesprocessF1(dataframe)
        self.assertEqual(result, True)

    def testprocess1ex3(self):
        dataframe = 'Weeks 18 - 21'
        result = tesprocessF1(dataframe)
        self.assertEqual(result, "No pattern matches")

    def testprocess1ex4(self):
        starttime = '00:00:00'
        teachingweeks = 13
        result = tesprocessF1_2(starttime,teachingweeks)
        self.assertEqual(result, "weeks day time hours")

    def testprocess1ex5(self):
        starttime = ''
        teachingweeks = 13
        result = tesprocessF1_2(starttime,teachingweeks)
        self.assertEqual(result, "weeks unscheduled hours")

    def testprocess1ex6(self):
        starttime = 19.0
        endtime = 19.0
        schedstart = 17.0
        result = tesprocesscalculationF1_3(starttime,endtime,schedstart)
        self.assertEqual(result, 0.5)

    def testprocess1ex7(self):
        starttime = 9.0
        endtime = 9.0
        schedstart = 7.0
        result = tesprocesscalculationF1_4(starttime,endtime,schedstart)
        self.assertEqual(result, 25.0)

    def testprocess2ex1(self):
        dataframe = 'Weeks 18-28'
        result = tesprocessF2(dataframe)
        self.assertEqual(result,True)
        
    def testprocess2ex2(self):
        dataframe = 'Term 2'
        result = tesprocessF2(dataframe)
        self.assertEqual(result, True)

    def testprocess2ex3(self):
        dataframe = 'Term 3'
        result = tesprocessF2(dataframe)
        self.assertEqual(result, True)

    @unittest.expectedFailure
    def testprocess1ex1fail(self):
        dataframe = 'Term 3'
        result = tesprocessF1(dataframe)
        self.assertEqual(result, True)

    @unittest.expectedFailure
    def testprocess2ex1fail(self):
        dataframe = 'Term 1'
        result = tesprocessF2(dataframe)
        self.assertEqual(result, True)
    
    @unittest.expectedFailure
    def testcounterfail(self):
        self.assertEqual(UnboundLocalError,countertest(-1))

    @unittest.expectedFailure
    def testcounterfail2(self):
        self.assertEqual(UnboundLocalError,countertest(str(abc)))
        
unittest.main()
