Function RandomFirstName() As String
    Dim firstNameArray() As Variant
    Dim randomIndex As Integer

    ' Predefined list of first names
    firstNameArray = Array("Alice", "Bob", "Charlie", "David", "Emma", "Frank", "Grace", "Henry", "Isabella", "Jack", _
                           "Kate", "Liam", "Mia", "Noah", "Olivia", "Peter", "Quinn", "Rachel", "Samuel", "Taylor", _
                           "Ursula", "Victor", "Wendy", "Xavier", "Yvonne", "Zachary")
    
    ' Calculate the number of first names in the array
    Dim numFirstNames As Integer
    numFirstNames = UBound(firstNameArray) - LBound(firstNameArray) + 1
    
    ' Generate a random index
    Randomize
    randomIndex = Int((numFirstNames * Rnd))
    
    ' Return the randomly selected first name
    RandomFirstName = firstNameArray(randomIndex)
End Function

Function RandomLastName() As String
    Dim lastNameArray() As Variant
    Dim randomIndex As Integer

    ' Predefined list of last names
    lastNameArray = Array("Smith", "Johnson", "Williams", "Jones", "Brown", "Davis", "Miller", "Wilson", "Moore", "Taylor", _
                          "Anderson", "Thomas", "Jackson", "White", "Harris", "Martin", "Thompson", "Garcia", "Martinez", "Robinson", _
                          "Clark", "Rodriguez", "Lewis", "Lee", "Walker")
    
    ' Calculate the number of last names in the array
    Dim numLastNames As Integer
    numLastNames = UBound(lastNameArray) - LBound(lastNameArray) + 1
    
    ' Generate a random index
    Randomize
    randomIndex = Int((numLastNames * Rnd))
    
    ' Return the randomly selected last name
    RandomLastName = lastNameArray(randomIndex)
End Function
Function GenerateRandomMeasurement(minValue As Double, maxValue As Double) As Double
    GenerateRandomMeasurement = WorksheetFunction.RandBetween(minValue * 10, maxValue * 10) / 10
End Function
Function GenerateRandomNumber(minValue As Integer, maxValue As Integer) As Double
    GenerateRandomMeasurement = WorksheetFunction.RandBetween(minValue * 1000000000, maxValue * 9999999999#)
End Function

Function GenerateRandomFemale() As Boolean
    GenerateRandomFemale = (Rnd() < 0.5) ' 50% chance of being true (female)
End Function

Function GenerateRandomPhoneNumber() As String
    Dim phoneNumber As String
    Dim randomNumber As Integer
    
    ' Generate random area code (3 digits)
    randomNumber = Int((999 - 100 + 1) * Rnd + 100)
    phoneNumber = phoneNumber & randomNumber
    
    ' Generate random exchange code (3 digits)
    randomNumber = Int((999 - 100 + 1) * Rnd + 100)
    phoneNumber = phoneNumber & randomNumber
    
    ' Generate random line number (4 digits)
    randomNumber = Int((9999 - 1000 + 1) * Rnd + 1000)
    phoneNumber = phoneNumber & randomNumber
    
    ' Return the generated phone number
    GenerateRandomPhoneNumber = phoneNumber
End Function


Sub NC_Experiment()
    
    Dim generatedFirstName As String
    Dim generatedLastName As String
    Dim randomPhoneNumber As String
    randomPhoneNumber = GenerateRandomPhoneNumber()

    
    ' Call the RandomFirstName function to get a random first name
    generatedFirstName = RandomFirstName()
    
    ' Call the RandomLastName function to get a random last name
    generatedLastName = RandomLastName()
    
    NCInput_Form.NC_RankInput.Value = "FSgt"
    NCInput_Form.NC_SurnameInput.Value = generatedLastName
    NCInput_Form.NC_FirstNameInput.Value = generatedFirstName
    NCInput_Form.NC_TelephoneInput.Value = randomPhoneNumber
    NCInput_Form.NC_EmailInput.Value = "Jesus@hotmail.com"
    NCInput_Form.NC_HeadInput.Value = GenerateRandomMeasurement(19, 26)
    NCInput_Form.NC_NeckInput.Value = GenerateRandomMeasurement(12.5, 20)
    NCInput_Form.NC_ChestInput.Value = GenerateRandomMeasurement(24, 64)
    NCInput_Form.NC_WaistInput.Value = GenerateRandomMeasurement(30, 63)
    NCInput_Form.NC_HipsInput.Value = GenerateRandomMeasurement(30, 68)
    NCInput_Form.NC_HeightInput.Value = GenerateRandomMeasurement(55, 76)
    NCInput_Form.NC_FootLInput.Value = GenerateRandomMeasurement(215, 330)
    NCInput_Form.NC_FootWInput.Value = GenerateRandomMeasurement(85, 130)
    NCInput_Form.NC_HandLInput.Value = GenerateRandomMeasurement(6, 10)
    NCInput_Form.NC_FemaleInput.Value = GenerateRandomFemale()
    
    NCInput_Form.Show
End Sub
Sub StressTest()
    Dim i As Integer
    
    i = 1
    Do While i <= 100
        NC_Experiment
        Call NCInput_Form.NC_SubmitButton_Click
        i = i + 1
    Loop
End Sub