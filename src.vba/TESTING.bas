Sub Testing()
    
    Dim s As Collection
    Set s = New Collection
    s.Add "25", "head"

    Dim o As String
    o = GetSize("Tilly", s)
    Debug.Print o
    
    s.Add "295", "FootL"
    s.Add "110", "FootW"
    o = GetSize("Leather Boots", s)
    Debug.Print o
    
    s.Add "14.2", "neck"
    s.Add "41.2", "chest"
    s.Add "55", "height"
    s.Add False, "IsMale"
    o = GetSize("Collar Shirt", s)
    Debug.Print o
    
    'Debug.Print s.Item("neck")
    'Debug.Print s.Item("NECK")
    'Sheets("Mia_Wilson_B1CF9FB7").Range("B6:B8") = Array(1, 2, 3)
End Sub
