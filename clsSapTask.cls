VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSapTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public status As String
Public owners As String

Private m_full_task As String
Private m_short_task As String
Private m_material As String
Private m_plant_so As String
Private m_task_type As String

Property Let full_task(ByVal value As String)
    m_full_task = value
End Property
Property Get full_task() As String
    full_task = m_full_task
End Property

Property Let short_task(ByVal value As String)
    m_short_task = value
End Property
Property Get short_task() As String
    short_task = m_short_task
End Property

Property Let material(ByVal value As String)
    m_material = value
End Property
Property Get material() As String
    material = m_material
End Property
Property Let plant_so(ByVal value As String)
    m_plant_so = value
End Property
Property Get plant_so() As String
    plant_so = m_plant_so
End Property

Property Let task_type(ByVal value As String)
    m_task_type = value
End Property
Property Get task_type() As String
    task_type = m_task_type
End Property


Sub split_full_task(ByVal full_task As String)
    m_full_task = UCase(Trim(full_task))
    m_short_task = Split(m_full_task, " REQUIRES ")(1)
    m_material = Split(Split(m_full_task, " MATERIAL ")(1), " ")(0)
    m_plant_so = Right(Split(m_full_task, " REQUIRES ")(0), 4)
    Dim utils As New clsUtils
    If utils.contains(m_full_task, "plant") Then
        m_task_type = "Plant"
    ElseIf utils.contains(m_full_task, "sales org") Then
        m_task_type = "Plant"
    Else
        m_task_type = "n/a"
    End If
End Sub

