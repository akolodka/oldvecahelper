VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCacheKeysNumName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const CACHE_HEADER = "numName"
Const KEY_LASTNUM = "lastNumber"

Const KEY_USE_EMPLOYEE_SECOND = "useEmpSecond"
Const KEY_USE_EMPLOYEE_THIRD = "useEmpThird"

Const KEY_USE_EMPLOYEE_OTHER = "useEmpOther"
Const KEY_EMPLOYEE_OTHER = "employeeOther"

Const KEY_SAVEASCOPY = "saveAsCopy"

Friend Property Get header() As String
    header = CACHE_HEADER
End Property

Friend Property Get cachedNumber() As String
    cachedNumber = KEY_LASTNUM
End Property

Friend Property Get employeeOther() As String
    employeeOther = KEY_EMPLOYEE_OTHER
End Property

Friend Property Get useEmployeeSecond() As String
    useEmployeeSecond = KEY_USE_EMPLOYEE_SECOND
End Property

Friend Property Get useEmployeeThird() As String
    useEmployeeThird = KEY_USE_EMPLOYEE_THIRD
End Property

Friend Property Get useEmployeeOther() As String
    useEmployeeOther = KEY_USE_EMPLOYEE_OTHER
End Property

Friend Property Get saveAsCopy() As String
    saveAsCopy = KEY_SAVEASCOPY
End Property

