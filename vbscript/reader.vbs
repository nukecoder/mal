Class Reader
      Public m_Tokens
      Public Position

      Public Property Get Peek
             if Position > m_Tokens.Count then
               Peek = trim(StripComas(m_Tokens(m_Tokens.Count - 1)))
             else 
               Peek = trim(StripCommas(m_Tokens(Position)))
             end if
      End Property

      Public Property Get GetNext
             GetNext = Peek
             Position = Position + 1
      End Property

      Public Property Set Tokens(tks)
             Set m_Tokens = tks
      End Property 

      Public Property Get Tokens()
             Set Tokens = m_Tokens
      End Property

   Public Sub PrintTokens()
      Dim x
      Wscript.StdOut.WriteLine("there are " & m_Tokens.Count & " tokens: ")
      Dim i
      For each x in m_Tokens
	 Wscript.StdOut.WriteLine(x.Value)
      Next
   End Sub

End Class

Function StripCommas(a_String)
   StripCommas = Replace(a_String, ",", "", 1, -1, 1)
End Function

Function tokenizer(str)
    dim str1
    set aR = New RegExp
    aR.IgnoreCase = True
    aR.Global = True
    aR.Pattern = "\r\n"
    str1 = aR.Replace(str, "\n")
    Set myRegExp = New RegExp
    myRegExp.IgnoreCase = True
    myRegExp.Global = True
    myRegExp.Pattern = "[\s,]*(~@|[\[\]{}()'`~^@]|" & Chr(34) & "(?:\\.|[^\\" & Chr(34) & "])*" & Chr(34) & "|;.*|[^\s\[\]{}('" & Chr(34) & "`,;)]*)"
    Set tokenizer = myRegExp.Execute(trim(str1))
End Function

Function read_str(myStr)
    Dim myReader
    Set myReader = new Reader
    Set myReader.Tokens = tokenizer(myStr)
' The tokenizer doesn't catch " by itself. So there is no way to validate a string
' as a string except before it passes through the tokenizer
' I using a naive check of the string.
    if VarType(myStr) = vbString then
      read_str = read_form(myReader)
    else
      read_str = "expected '" & Chr(34) & "', got EOF"
    end If
End Function

Function read_form(rdr)
    Dim token
    token = trim(rdr.peek())
    Select Case Left(trim(token),1)
      Case "("
           read_form = read_list(rdr)
      Case "["
           read_form = read_vector(rdr)
      Case "{"
           read_form = read_map(rdr)
      Case "'"
           read_form = read_quote(rdr)
      Case "`"
           read_form = read_quasiquote(rdr)
      Case "~"
           if Left(trim(token), 2) = "~@" then
             read_form = read_splice_unquote(rdr)
           else 
             read_form = read_unquote(rdr)
           end if
      Case "^"
           read_form = read_with_meta(rdr)
      Case "@"
           read_form = read_deref(rdr)
      Case else
           read_form = read_atom(rdr)
    End Select
End Function

Function read_seq(rdr, the_end, the_tag)
  Dim result(), token, Serr, form
  ReDim result(0)
  result(0) = Null
  serr = ""
  token = rdr.getNext() 'move position past opening character
  token = trim(rdr.peek())
  Do while trim(token) <> the_end
    form = read_form(rdr)
    if typeName(form) = "string" then
      Serr = form
      Exit Do
    End if
    if isNull(result(0)) then
      result(0) = form
    else
      redim preserve result(UBound(result) + 1)
      result(UBound(result)) = form
    end if
    if token = "" Then 
      Serr = "expected '" & the_end & "', got EOF"
      Exit Do
    end if
    token = rdr.getNext()
    token = rdr.peek()
  Loop
  if Serr <> "" Then
     read_seq = Serr
  else 
    read_seq = Array(the_tag, result)
  end if
End Function

Function read_list(rdr)
  read_list = read_seq(rdr, ")", "list")
End Function

Function read_vector(rdr)
  read_vector = read_seq(rdr, "]", "vector")
End Function 

Function read_map(rdr)
  read_map = read_seq(rdr, "}", "hashmap")
End Function 

Function read_atom(rdr)
  Dim token
  token = rdr.peek()
  If IsNumeric(token) then
      read_atom = Array("integer", read_number(token))
  elseif Left(trim(token),1) = Chr(34) and Right(trim(token),1) = Chr(34) then
      read_atom = Array("string", read_string(token))
  elseif token = "true" then
      read_atom = Array("true", token)
  elseif token = "false" then
      read_atom = Array("false", token)
  elseif token = "nil" then
      read_atom = Array("nil", token)
  elseif Left(trim(token), 1) = ":" then
      read_atom = Array("keyword", trim(token))
  else
      read_atom = Array("symbol", token)
  End if
End Function

Function read_number(str)
   if InStr(str, ".") > 0 then
      read_number = CDbl(str)
   else
      read_number = CInt(str)
   end if
End Function

Function read_string(str)
  if trim(str) <> Chr(34) & Chr(34) Then 
  ' Remove the leading and trailing "
  str = Left(str,len(str) - 1)
  str = Right(str,len(str)-1)

  ' \" -> "
  str = Replace(str, "\" & Chr(34), Chr(34))

  ' \n -> vbLf
  str = Replace(str, "\n", vbLf)

  ' \\ -> \
  str = Replace(str, "\\", "\")
  End If 

  read_string = trim(str)

End Function

Function read_a_quote(rdr, atype)
  rdr.getNext() ' move past '
  read_a_quote = Array(atype, pr_str(read_form(rdr), False))
End Function

Function read_quote(rdr)
  read_quote = read_a_quote(rdr, "quote")
End Function

Function read_quasiquote(rdr)
  read_quasiquote = read_a_quote(rdr, "quasiquote")
End Function

Function read_unquote(rdr)
  read_unquote = read_a_quote(rdr, "unquote")
End Function

Function read_splice_unquote(rdr)
  read_splice_unquote = read_a_quote(rdr, "splice-unquote")
End Function

Function read_with_meta(rdr)
  Dim meta
  rdr.getNext ' move past '^'
  meta = read_map(rdr) ' get the metadata as a hashmap
  rdr.getNext ' move past closing brace of hashmap
  read_with_meta = Array("list", Array(Array("symbol", "with-meta"), read_form(rdr), meta))
End Function

Function read_deref(rdr)
  Dim symbol
  rdr.getNext ' move past '@'
  symbol = read_form(rdr)
  read_deref = Array("list", Array(Array("symbol", "deref"), symbol))
End Function