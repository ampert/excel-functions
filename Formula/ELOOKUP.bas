Function ELOOKUP(lookup_value As Range, lookup_range As Range, value_range As Range) As Variant

    ELOOKUP = Application.WorksheetFunction.Index(value_range, _
        Application.WorksheetFunction.Match(lookup_value, lookup_range, 0))

End Function
