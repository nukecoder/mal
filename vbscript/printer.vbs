Function pr_str(ast, print_readably)
  if typeName(ast) = "Variant()" Then
    Select Case ast(0)
      Case "integer"
           pr_str = CStr(ast(1))
      Case "symbol"
           pr_str = ast(1)
      Case "list"
           pr_str = pr_seq(ast(1), print_readably, "(", ")")
      Case "vector"
           pr_str = pr_seq(ast(1), print_readably, "[", "]")
      Case "hashmap"
           pr_str = pr_seq(ast(1), print_readably, "{", "}")
      Case "string"
           pr_str = pr_string(ast(1), print_readably)
      Case "true"
           pr_str = "true"
      Case "false"
           pr_str = "false"
      Case "nil"
           pr_str = "nil"
      Case "quote"
           pr_str = "(quote " & ast(1) & ")"
      Case "quasiquote"
           pr_str = "(quasiquote " & ast(1) & ")"
      Case "unquote"
           pr_str = "(unquote " & ast(1) & ")"
      Case "splice-unquote"
           pr_str = "(splice-unquote " & ast(1) & ")"
      Case "keyword"
           pr_str = ast(1)
      Case else
           pr_str = ast(1)
    End Select
  else
    pr_str = ast
  end if
End Function

Function pr_seq(l, print_readably, start, tend)
  Dim x, result
  result = ""
  if isNull(l(0)) then 'empty list
    result = ""
  else
    For x = 0 to UBound(l) - 1
      result = result & pr_str(l(x), print_readably) & " "
    Next
    result = result & pr_str(l(UBound(l)), print_readably)
  end if
  pr_seq = start & trim(result) & tend
End Function

Function pr_string(s, print_readably)
  if s <> Chr(34) & Chr(34) then
    if print_readably then
      ' \\ -> \
      s = Replace(s, "\", "\\")

      ' \n -> vbLf
      s = Replace(s, vbLf, "\n")

      ' " -> \" 
      s = Replace(s, Chr(34), "\" & Chr(34), 1, len(s) - 1)
    end if
    pr_string = Chr(34) & s & Chr(34)
  else
    pr_string = s
  end if

End Function

