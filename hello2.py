# bank implements a simple bank account class and the user class will be created and for 
# each user the bank account will be created.
# 
import random
import xlwt, xlrd

class User:
    def __init__(self, name, balance):
        self.name = name
        self.balance = balance

        self.account_number = random.randint(100000,9999999)
        self.credit_card = random.randint(10000,99999)
        self.pan_card = random.randint(100000,999999) 

    def __str__(self): 
        return f"{self.name} has {self.balance} in their account and their account number is {self.account_number}and their credit card number is {self.credit_card} and their pan card number is {self.pan_card}"

class Bank:
    def __init__(self):
        self.users = []
        self.account = []
        

    def addUser(self, user):
        self.users.append(user)
        self.account.append(user.account_number)
         

    def printAllUsers(self):
        for user in self.users:
            print(user)
        for account in self.account:
            print(account)     

    def writeUsersToExcel(self):
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Users")
        row = 0
        for user in self.users:
            ws.write(row, 0, user.name)
            ws.write(row, 1, user.balance)
            ws.write(row, 2, user.account_number)
        

            row += 1
        wb.save("users.xls")

    def readUsersFromExcel(self):
        wb = xlrd.open_workbook("users.xls")
        ws = wb.sheet_by_name("Users")
        for row in range(ws.nrows):
            name = ws.cell(row, 0).value
            balance = ws.cell(row, 1).value
            account_number = ws.cell(row, 2).value
            user = User(name, balance)
             
            self.users.append(user)



    def getUser(self, name):
        for user in self.users:
            if user.name == name:
                return user
        return None



    def profile(self, name):
        user = self.getUser(name)
        if user:
            print("Name: ", user.name)
            print("Balance: ", user.balance)
            print("Account number ",user.account_number)
            # create a file for the user and write the name and balance in the file.
            wb = xlwt.Workbook()
            ws = wb.add_sheet("Profile")
            ws.write(0, 0, user.name)
            ws.write(0, 1, user.balance)
            ws.write(0,2,user.account_number)
            wb.save(user.name + ".xls")
        else:
            print("User not found")

    def deposit(self, name, amount):
        user = self.getUser(name)
        if user:
            user.balance += amount
          
            wb = xlwt.Workbook()
            ws = wb.add_sheet("Profile")
            ws.write(0, 0, user.name)
            ws.write(0, 1, user.balance)
            ws.write(0,2,user.account_number)
            wb.save(user.name + ".xls")
        else:
            print("User not found")


bank = Bank()
bank.addUser(User("John", 100))
bank.addUser(User("Mary", 200))
bank.addUser(User("Peter", 300))
bank.readUsersFromExcel()
bank.writeUsersToExcel()
bank.deposit("Mary", 100)
# bank.profile("Mary")

bank.printAllUsers()
 



