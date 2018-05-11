Attribute VB_Name = "Module1"

Sub DurationToDays()

Dim task As task
Dim temp As String

    For Each task In ActiveProject.Tasks
        temp = DurationFormat(task.Duration, pjDays)
        task.Duration = temp
    Next task
End Sub
