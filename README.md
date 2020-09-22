<div align="center">

## Dynamically Pass Form Variables through Forms


</div>

### Description

This code passes all form variables by creating hidden copies of every form variable that was passed to it. This is great for gathering information over multiples forms (i.e. Order Processing).
 
### More Info
 
Active Server Pages (Tested on with IIS 4.0)

A string representing the hidden form varibles.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Shell](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-shell.md)
**Level**          |Beginner
**User Rating**    |3.8 (19 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__4-3.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-shell-dynamically-pass-form-variables-through-forms__4-6467/archive/master.zip)





### Source Code

```
'This is the call here....
'Response.Write PassFormVariables
Function PassFormVariables()
Const INPUT_START = "<input type=hidden name="
Const INPUT_MID = " value="
Const INPUT_END = ">" & Chr(13) & Chr(10)
Dim var_name
Dim var_value
	PassFormVariables = ""
	if len(request.form) > 0 then
		for each var_name in request.form
			var_value = request.form(var_name)
			PassFormVariables = PassFormVariables & _
				INPUT_START & var_name & INPUT_MID & var_value & INPUT_END
		next
	end if
End Function
```

