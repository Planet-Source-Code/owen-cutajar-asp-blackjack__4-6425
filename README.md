<div align="center">

## ASP Blackjack


</div>

### Description

Play Blackjack against the computer
 
### More Info
 
Code should be fully functional. One assumption is that card generation is totally random. There is no check that a card has already been extracted from the deck in the same game.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Owen Cutajar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/owen-cutajar.md)
**Level**          |Intermediate
**User Rating**    |3.3 (20 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Games](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/games__4-13.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/owen-cutajar-asp-blackjack__4-6425/archive/master.zip)





### Source Code

```
<%
'--------------------------------------------------------------
' ASPBlackJack v1.0
' (x) Ugh!! 2000 - 15/10/2000
'
' Converted to ASP from JScript obtained at
' http://www.javascriptsearch.com/scripts/Games/blackjack.html
' Original author: Unknown
'
' Written for http://www.only-network.com/games Any comments,
' flames, requests, postcards etc to owen@cutajar.net
'--------------------------------------------------------------
Option Explicit
'------------------------------------------------
' History Routines
'------------------------------------------------
sub ClearHistory
	Dim q
	For q = 1 to 10
		Session("DHistory" & q) = ""
		Session("UHistory" & q) = ""
	next
end sub
sub WriteHistory(Entity,Message)
	Dim q
	q = 1
	While Session( Left(Entity,1) & "History" & q) <> ""
		q = q + 1
	Wend
	Session( Left(Entity,1) & "History" & q) = Message
end sub
Function ReadHistory(Entity)
	Dim returnValue,q,currentmessage
	returnvalue = ""
	for q = 1 to 10
		currentmessage = Session( Left(Entity,1) & "History" & q)
		If currentmessage <> "" Then
			returnvalue = returnvalue & "- " & currentmessage & "<br>"
		end if
	next
	ReadHistory = returnValue
end Function
'------------------------------------------------
' Select a random card
'------------------------------------------------
function pickCard()
	Dim iRandom
	Randomize()
	iRandom = Int(13 * Rnd + 1)
	pickCard = iRandom
end function
'------------------------------------------------
' Select a random suit
'------------------------------------------------
function pickSuit()
	Dim iRandom
	Randomize()
	iRandom = Int(4 * Rnd + 1)
	if iRandom = 1 then
	   pickSuit = "Spades"
	elseif iRandom = 2 then
	   pickSuit = "Clubs"
	elseif iRandom = 3 then
	   pickSuit = "Diamonds"
	else
	   pickSuit = "Hearts"
	end if
end function
'------------------------------------------------
' Translate card number to words
'------------------------------------------------
function cardName(card)
	if card = 1 then
		cardName = "Ace"
	elseif card = 11 then
		cardName = "Jack"
	elseif card = 12 then
		cardName = "Queen"
	elseif card = 13 then
		cardName = "King"
	else
		cardName = card
	end if
end function
'------------------------------------------------
' Work out value of card
'------------------------------------------------
function cardValue(card)
  	if card = 1 then
   		cardValue = 11
	elseif card > 10 Then
		cardValue = 10
	else
		cardValue = card
	end if
end function
'------------------------------------------------
' Work out value of card
'------------------------------------------------
function selectCard(strWho)
	Dim iCardNum,theCard
	iCardNum = pickCard
	theCard = CardName(iCardNum) & " of " & pickSuit
	Response.Write strWho & " picked a " & theCard & "<br>"
	selectCard = CardValue(iCardNum)
	WriteHistory strWho , theCard
end function
'------------------------------------------------
' Show current status on screen
'------------------------------------------------
sub ReportStatus
	Response.Write "<p>"
	Response.Write "<table border=3><tr><td>Player</td><td>Hand</td><td>Score</td></tr>"
	Response.Write "<tr><td>You</td><td>" & ReadHistory("User") & "</td>"
	Response.Write "<td>" & Session("User") & "</td></tr>"
	Response.Write "<tr><td>Dealer</td><td>" & ReadHistory("Dealer") & "</td>"
	Response.Write "<td>" & Session("Dealer") & "</td></tr></table>"
end sub
'------------------------------------------------
' Work out value of card
'------------------------------------------------
sub NewHand
	ClearHistory
  	Session("Dealer") = 0
  	Session("User") = 0
  	Session("Dealer") = Session("Dealer") + SelectCard("Dealer")
  	Session("User") = Session("User") + SelectCard("User")
	Session("GameOver") = False
	ReportStatus
end Sub
'------------------------------------------------
' Dealer Move
'------------------------------------------------
sub DealerMove
  	while Session("Dealer") < 17
   		Session("Dealer") = Session("Dealer") + SelectCard("Dealer")
  	wend
	ReportStatus
end sub
'------------------------------------------------
' You Move
'------------------------------------------------
sub UserMove
  	Session("User") = Session("User") + SelectCard("User")
  	if Session("User") > 21 Then
     	Response.Write "<h2>You busted!</h2>"
		Session("GameOver") = True
  	end if
	ReportStatus
end sub
'------------------------------------------------
' Get Game Status
'------------------------------------------------
Sub LookAtHands
  	if Session("Dealer") > 21 Then
    		Response.Write "<h2>House busts! You win!</H2>"
  	elseif Session("User") > Session("Dealer") Then
    		Response.Write "<h2>You win</H2>"
  	elseif Session("User") = Session("Dealer") Then
		Response.Write "<h2>Push!</H2>"
  	else
   		Response.Write "<h2>House wins!</H2>"
	end if
	Session("GameOver") = True
End Sub
' Maincode starts here
Dim action
Response.Write "<center>"
action = Request("action") & ""
If action = "Hit Me" Then
	UserMove
Elseif action = "Stand" Then
	DealerMove
	LookAtHands
Else
	NewHand
End If
Response.Write "<form action=blacksource.asp>"
If Session("GameOver") <> True Then
	Response.Write "<INPUT TYPE=SUBMIT NAME=action VALUE='Hit Me'>"
	Response.Write "<INPUT TYPE=SUBMIT NAME=action VALUE='Stand'>"
End if
Response.Write "<INPUT TYPE=SUBMIT NAME=action VALUE='New Hand'>"
Response.Write "</FORM>"
Response.Write "</center>"
%>
```

