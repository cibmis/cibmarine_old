<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/email.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_email_STRING
  MM_editTable = "email"
  MM_editRedirectUrl = "thankyou.asp"
  MM_fieldsStr  = "FirstName|value|LastName|value|Street_Address|value|City|value|State|value|Zip|value|PhoneNumber|value|EmailAddress|value"
  MM_columnsStr = "FirstName|',none,''|LastName|',none,''|[Street Address]|',none,''|City|',none,''|State|',none,''|Zip|',none,''|PhoneNumber|',none,''|EmailAddress|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>CIB Marine Bancshares, Inc.</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link rel="stylesheet" href="../css/cibstyles.css" type="text/css" />

<meta name="description" content="CIB Marine Bancshares, Inc.">
<meta name="keywords" content="best, cd rates,cds,yield,yields,highest,one,rates,monitor,financial,cd,markets,bank,banking,safety,certificate,deposit,variable,index,indexed,historical,investment,best,money-rates,savings,newsletter,advice,FDIC,
jumbo

mortgage rates, CD rates, interest rates, credit cards, mortgages, home equity loans, certificate of deposit, calculators, banks, checking, interest rates, money market funds

Wisconsin, Illinois, Nebraska, Florida, Indianapolis, Indy, Arizona, Milwaukee, Chicago, Miami, Champaign, Scottsdale, Omaha, Nevada

Holding, Corporation, Corp, National, Bank, Banks, Banking, Louisiana, LA, Mississippi, MS, Alabama, AL, Florida, FL, Financial, FDIC, Equal, Housing, Lender, Securities, Service, HB, Online, family, industry, Security, Corporate, Profile, community, Career, job, jobs, openings, positions, postings, opportunities, Resume, recruiting, Privacy, Disclosure, disclaimers, brokerage, Checking, deposit, account, Select, Save, overdraft protection, 
checks, charge, student, senior, Check, Card, credit, debit, atm, Visa, savings, accounts, interest, money, market, withdrawal, interest, rate, Fixed, certificates, deposits, CD, Retirement, tax, benefits, IRA, individual, Roth, investment, contributions, 401k, 401-k, 401(k), education, educational, Trust, Planning, assets, Loans, borrowing, equity, consumer, car, auto, personal, line of credit, LOC, loan, secured, automobile, payment, mezzanine, lending, lend, mortgage, ARM, adjustable rate, FHA, VA,  financing, refinance, pension, Business, purchase, SEP, mutual, funds, mutual funds, risk, bonds, online banking, savings bonds, certificate of deposit, internet banking, money market, checking account, commercial banking, investments, mortgage, Internet banking, Southeast, northeast, midwest, northeast, west, western, southwest, south, money market account, certificate of deposit rate, bank rate, banks, internet bill pay, internet bill payment">

</head>
<body link="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="../art/background.gif" vlink="#9C0000" text="#000000">
<table bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" width="100%">
  <!-- fwtable fwsrc="CIB Design 2.png" fwbase="index.jpg" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="1" -->
  <tr>
    <td colspan="3" bgcolor="#323264"> 
      <table bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" width="90%">
        <tr bgcolor="#9C0000"> 
          <td width="486" align="left" bgcolor="#323264"><img src="../art/clipper_ship_head.jpg" alt="CIB Marine Bancshares, Inc. Header" name="img_header" width="528" height="49" border="0" id="img_header" /></td>
	      <td align="right" valign="top" width="354" bgcolor="#323264"> 
            <table width="261" border="0" cellspacing="0" cellpadding="0" height="19">
              <tr> 
                <td width="261"> 
                  <div class="white_text" align="right"><a href="../privacy.asp">Privacy Statement</a> | <a href="../contact/contactus.asp">Contact Us</a> | <a href="../index.asp">Home</a></div>
                </td>
                <td width="10"> 
                  <div align="right"><img src="../art/spacer.gif" width="20" height="12" border="0" /></div>
                </td>
              </tr>
            </table>
          </td>
	  </tr>
	</table>
    </td>
  </tr>
  <tr>
    <td colspan="3">
      <table width="810" border="0" cellspacing="0" cellpadding="0" height="24">
        <tr> 
          <td width="12"><img src="../art/spacer.gif" width="10" height="12" border="0" /></td>
          <td class="cibBreadcrumb" width="189">        
             <%= FormatDateTime(now,1) %>
              
          </td>
          <td width="609"><a href="../index.asp"></a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
   <td colspan="3">
      <table bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" width="786">
        <tr> 
          <td width="201" bgcolor="#A3A3A3" valign="top"> 
            <table bgcolor="#A3A3A3" border="0" cellpadding="0" cellspacing="0" width="76">
              <tr> 
                <td bgcolor="#A3A3A3" valign="top">
<table width="201" border="0" cellspacing="0" cellpadding="10" height="145">
                    <tr> 
                      <td bgcolor="#A3A3A3" valign="top"><img src="../art/life_at_CIB.gif" alt="Life At CIB Marine" width="145" height="9" /><br />
                        <img src="../art/spacer.gif" width="150" height="6" border="0" /><br />
                        <table class="white_nav" width="169" border="0" cellspacing="0" cellpadding="4">
                          <tr>
                            <td width="4"><img src="../art/icon_wht_bullet.gif" width="4" height="4" /></td>
                            <td width="146" nowrap="nowrap"><a href="../corp_info/shareholders.asp">Investor Relations</span></a></td>
                          </tr>
                          <tr>
                            <td width="4"><img src="../art/icon_wht_bullet.gif" width="4" height="4" /></td>
                            <td width="146" nowrap="nowrap"><a href="../corp_info/banking_subs.asp">Banking Subsidiaries</a></td>
                          </tr><!--                          <tr>
                            <td width="4"><img src="art/icon_wht_bullet.gif" width="4" height="4" /></td>
                            <td width="146" nowrap="nowrap"><a href="corp_info/nonbanking_subs.asp">
							Non-Banking Subsidiaries/a></td>
                          </tr>-->
                          <tr>
                            <td><img src="../art/icon_wht_bullet.gif" width="4" height="4" /></td>
                            <td nowrap="nowrap"><a href="../corp_info/careers.asp">Careers</a></td>
                          </tr>
                          <tr>
                            <td><img src="../art/icon_wht_bullet.gif" width="4" height="4" /></td>
                            <td nowrap="nowrap"><a href="../documents/codeofethicspolicy_final.pdf" target="_blank">Code of Ethics <br />
                            Policy [PDF]</a></td>
                          </tr>
                    </table>                    </td>
                    </tr>
                    <tr> 
                      <td bgcolor="#A3A3A3" valign="top"><img src="../art/career_ops.gif" alt="Career Opportunities" width="145" height="9" /><br />
                        <img src="../art/spacer.gif" width="150" height="6" border="0" />
                        <table class="white_nav" width="169" border="0" cellspacing="0" cellpadding="4">
                          <tr>
                            <td width="4"><img src="../art/icon_wht_bullet.gif" width="4" height="4" /></td>
                            <td width="146" nowrap="nowrap"><a href="../financial_info/sec_filings.asp">SEC Filings</a></td>
                          </tr>
                          <tr>
                            <td width="4"><img src="../art/icon_wht_bullet.gif" width="4" height="4" /></td>
                            <td width="146" nowrap="nowrap"><a href="../financial_info/section16.asp">Section 16 Filings</a></td>
                          </tr>
                          <!--                          <tr>
                            <td width="4"><img src="art/icon_wht_bullet.gif" width="4" height="4" /></td>
                            <td width="146" nowrap="nowrap"><a href="corp_info/nonbanking_subs.asp">
							Non-Banking Subsidiaries/a></td>
                          </tr>-->
                          <tr>
                            <td><img src="../art/icon_wht_bullet.gif" width="4" height="4" /></td>
                            <td nowrap="nowrap"><a href="../financial_info/press_releases.asp">Press Releases</a></td>
                          </tr>
                          <tr>
                            <td><img src="../art/icon_wht_bullet.gif" width="4" height="4" /></td>
                            <td nowrap="nowrap"><a href="../financial_info/holderletters.asp">Shareholder Letters</a></td>
                          </tr>
                        </table>
                        </td>
                    </tr>
                    <tr> 
                      <td bgcolor="#A3A3A3" valign="top">&nbsp;</td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
          <td valign="top" width="585"> 
            <table width="608" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td bgcolor="#292929" width="392"><img src="../images/heads/cibm.gif" alt="CIB Marine Bancshares, Inc." width="315" height="43" /></td>
                <td bgcolor="#292929" width="216">&nbsp;</td>
              </tr>
            </table>
            <img src="../images/heads/blank.gif" alt="Privacy Statement" width="608" height="26" /> 
            <table class="cibHome" width="585" border="0" cellspacing="0" cellpadding="15">
              <tr valign="top"> 
                <td height="51"> <p align="center"><strong><font size="3">Request 
                    for Shareholder Contact Information</font></strong></p>
                  <hr />
                  <p>To ensure that we have your most current contact information, 
                    please provide us with your name, address, phone number(s), and 
                    email address. </p>
<p><em>This information will be used solely for 
                        shareholder-related correspondence from CIB Marine or Computershare 
                        Investor Services, our stock transfer agent. </em> </p>
<form action="<%=MM_editAction%>" method="post" name="form1">
  <table align="center">
    <tr valign="baseline"> 
      <td nowrap align="right">First Name:</td>
      <td> <input type="text" name="FirstName" value="" size="32"> </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right">Last Name:</td>
      <td> <input type="text" name="LastName" value="" size="32"> </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right">Street Address:</td>
      <td> <input type="text" name="Street_Address" value="" size="32"> </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right">City:</td>
      <td> <input type="text" name="City" value="" size="32"> </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right">State:</td>
      <td> <input type="text" name="State" value="" size="32"> </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right">Zip:</td>
      <td> <input type="text" name="Zip" value="" size="32"> </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right">Phone Number:</td>
      <td> <input type="text" name="PhoneNumber" value="" size="32"> </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right">Email Address:</td>
      <td> <input type="text" name="EmailAddress" value="" size="32"> </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right">&nbsp;</td>
      <td> <input type="submit" value="Insert Record"> </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
</form>
				
				
				
				
				
				</td>
              </tr>
            </table>
            <p>&nbsp;</p>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td colspan="3" align="right" bgcolor="#000000"><img src="../art/img_footer.gif" alt="Copyright 2003 CIB Marine Bancshares, Inc." name="img_footer" width="350" height="23" border="0" id="img_footer" /></td>
  </tr>
  <tr>
    <td width="20%"><div align="center"><img src="../art/icons_member_assoc.gif" alt="Member FDIC and Equal Housing Lender logos" width="83" height="26" /></div></td>
    <td width="79%" class="grey_text">&nbsp;</td>
    <td width="1%"></td>
  </tr>
</table>
<p style="margin-bottom: 0;">&nbsp;</p>
</body>
</html>
