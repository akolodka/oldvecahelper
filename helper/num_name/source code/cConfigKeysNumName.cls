VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConfigKeysNumName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const CONFIG_HEADER = "numName"

Const KEY_NUM_DEPARTMENT = "numberDepartment"
Const KEY_NUM_LABORATORY = "numberLaboratory"
Const KEY_PATH_JOURNAL_VERIFICATION = "journalVerificationPath"

Const KEY_EMPLOYEE_MAJOR = "employeeMajor"
Const KEY_EMPLOYEE_SECOND = "employeeSecond"
Const KEY_EMPLOYEE_THIRD = "employeeThird"


Friend Property Get header() As String
    header = CONFIG_HEADER
End Property

Friend Property Get departmentNumber() As String
    departmentNumber = KEY_NUM_DEPARTMENT
End Property

Friend Property Get laboratoryNumber() As String
    laboratoryNumber = KEY_NUM_LABORATORY
End Property

Friend Property Get journalPath() As String
    journalPath = KEY_PATH_JOURNAL_VERIFICATION
End Property

Friend Property Get employeeMajor() As String
    employeeMajor = KEY_EMPLOYEE_MAJOR
End Property

Friend Property Get employeeSecond() As String
    employeeSecond = KEY_EMPLOYEE_SECOND
End Property

Friend Property Get employeeThird() As String
    employeeThird = KEY_EMPLOYEE_THIRD
End Property
