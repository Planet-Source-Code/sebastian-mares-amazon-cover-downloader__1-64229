VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'################################ ACD CORE #####################################
'###  License:          BSD (Berkeley Software Distribution) License         ###
'###  Copyright:        2006, MaresWEB. All rights reserved.                 ###
'###############################################################################
'###  Version:          1.1d                                                 ###
'###  Release Date:     2006-02-07T18:45:00+01:00                            ###
'###############################################################################

'############################### CHANGE LOG ####################################
'### Version 1.0 (2005-08-27T17:30:00+02:00):                                ###
'###    - First build                                                        ###
'### Version 1.1 (2005-09-06T17:45:00+02:00):                                ###
'###    - Added ability to limit the number of pages to process              ###
'### Version 1.1a (2005-09-07T10:45:00+02:00):                               ###
'###    - Minor bug fixes                                                    ###
'### Version 1.1b (2005-11-09T17:45:00+01:00):                               ###
'###    - Switch from DOMDocument40 to DOMDocument60                         ###
'###    - Switch from AWS 2005-07-26 to AWS 2005-10-05                       ###
'### Version 1.1c (2006-02-04T17:00:00+01:00):                               ###
'###    - Switch back from DOMDocument60 to DOMDocument40                    ###
'###    - Switch from AWS 2005-07-26 to AWS 2005-10-05                       ###
'### Version 1.1d (2006-02-07T18:45:00+01:00):                               ###
'###    - License changed from LGPL to BSD                                   ###
'###############################################################################

'################################# LICENSE #####################################
'### © 2006, MaresWEB. All rights reserved.                                  ###
'###                                                                         ###
'### Redistribution and use in source and binary forms,                      ###
'### with or without modification, are permitted provided                    ###
'### that the following conditions are met:                                  ###
'###                                                                         ###
'###    1. Redistributions of source code must retain the                    ###
'###       above copyright notice, this list of conditions                   ###
'###       and the following disclaimer.                                     ###
'###    2. Redistributions in binary form must reproduce the                 ###
'###       above copyright notice, this list of conditions and               ###
'###       the following disclaimer in the documentation and / or            ###
'###       other materials provided with the distribution.                   ###
'###                                                                         ###
'### This software is provided by MaresWEB "as is" and any express           ###
'### or implied warranties, including, but not limited to, the implied       ###
'### warranties of merchantability and fitness for a particular purpose      ###
'### are disclaimed. In no event shall MaresWEB be liable for any direct,    ###
'### indirect, incidental, special, exemplary, or consequential damages      ###
'### (including, but not limited to, procurement of substitute goods or      ###
'### services; loss of use, data, or profits; or business interruption)      ###
'### however caused and on any theory of liability, whether in contract,     ###
'### strict liability, or tort (including negligence or otherwise) arising   ###
'### in any way out of the use of this software, even if advised of the      ###
'### possibility of such damage.                                             ###
'###############################################################################

Option Explicit
DefLng A-Z

'Return-Code constants
Private Const ACD_OK As Long = 200                      'OK - Operation finished flawlessly
Private Const ACD_NOXMLPARSER As Long = 400             'The XML parser cannot be used
Private Const ACD_NOLOADXML As Long = 401               'The XML parser could not load the document
Private Const ACD_AWSERROR As Long = 402                'AWS reported an error
Private Const ACD_NOVALIDITYCHECK As Long = 403         'Unable to validate query
Private Const ACD_NOITEMCOUNT As Long = 404             'Unable to retrieve the number of items available
Private Const ACD_NORETRIEVEITEM As Long = 405          'Unable to select "Item" node - no items available
Private Const ACD_NOSMALLIMAGE As Long = 406            'Matches found, but none contains a small image
Private Const ACD_NOPAGECOUNT As Long = 407             'Unable to retrieve the number of pages available
Private Const ACD_QUERYRUNTIMEERROR As Long = 500       'Runtime error during query process
Private Const ACD_PARSERUNTIMEERROR As Long = 501       'Runtime error during parse process
Private Const ACD_VALIDATERUNTIMEERROR As Long = 502    'Runtime error during validation process
Private Const ACD_USERSTOP As Long = 600                'User triggered halt
Private Const AWS_ITEMSPERPAGE As Long = 10             'Number of items per AWS page
Private Const QS_ALLINPUT As Long = &HFF

Public Type AWSError
    strCode As String
    strMessage As String
End Type
Public Type AWSItem
    lngLargeImageHeight As Long
    lngLargeImageWidth As Long
    lngMediumImageHeight As Long
    lngMediumImageWidth As Long
    lngSmallImageHeight As Long
    lngSmallImageWidth As Long
    strArtist As String
    strASIN As String
    strLargeImageURL As String
    strMediumImageURL As String
    strProductGroup As String
    strSmallImageURL As String
    strTitle As String
End Type
Public Type RuntimeError
    lngLine As Long
    lngNumber As Long
    strDescription As String
    strSource As String
End Type

Private Declare Function GetQueueStatus Lib "user32" (ByVal lngFlags As Long) As Boolean
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef anyDestination As Any, ByVal lngLength As Long)

Private blnCancel As Boolean
Private blnXMLLoaded As Boolean
Private blnSmallImageAvailable As Boolean
Private WithEvents objXMLParser As DOMDocument40
Attribute objXMLParser.VB_VarHelpID = -1
Private udtAWSError As AWSError
Private udtAWSItems() As AWSItem
Private udtRuntimeError As RuntimeError

Public Event Done(ByVal lngStatusCode As Long)
Public Event Progress(ByVal lngCurrentPage As Long, ByVal lngTotalNumberOfPages As Long)
Public Event ProgressDataAvailable()
Public Event WorkStarted()

Public Property Get AWSError() As AWSError

1         AWSError = udtAWSError

End Property

Public Sub Cancel()

1         blnCancel = True

End Sub

Private Sub Class_Initialize()

          'Instantiate and configure XML parser
1         If objXMLParser Is Nothing Then
2             Set objXMLParser = New DOMDocument40
3             objXMLParser.setProperty "SelectionNamespaces", "xmlns:aws='http://webservices.amazon.com/AWSECommerceService/2005-10-05'"
4         End If

End Sub

Private Sub Class_Terminate()

          'Destroy XML parser
1         If Not objXMLParser Is Nothing Then
2             Set objXMLParser = Nothing
3         End If

End Sub

Public Property Get Item(ByVal lngItem As Long) As AWSItem

        Dim udtDummyAWSItem As AWSItem

1         If Not lngItem = 0 Then
2             Item = udtAWSItems(Abs(lngItem) - 1)
3           Else
4             udtDummyAWSItem.lngLargeImageHeight = UBound(udtAWSItems) + 1
5             Item = udtDummyAWSItem
6         End If

End Property

Private Sub objXMLParser_OnReadyStateChange()

1         If Not objXMLParser.readyState = 4 Then
2             blnXMLLoaded = False
3           Else
4             blnXMLLoaded = True
5         End If

End Sub

Private Function ParseResponse(ByVal lngCurrentPage As Long) As Long

        Dim lngArrayIndex As Long
        Dim lngCurrentItem As Long
        Dim lngTotalNumberOfItems As Long
        Dim objNode As IXMLDOMNode
        Dim objNodes As IXMLDOMNodeList

1         On Error GoTo ErrorHandler

          'Select all items and stop on failure
2         Set objNodes = objXMLParser.selectNodes("aws:ItemSearchResponse/aws:Items/aws:Item")
3         If objNodes Is Nothing Then
4             ParseResponse = ACD_NORETRIEVEITEM
5             Exit Function
6         End If

7         lngTotalNumberOfItems = objNodes.length - 1

          'Loop through all items
8         Do
              'Check if there are queued message and parse them
9             If GetQueueStatus(QS_ALLINPUT) Then
10                DoEvents
11            End If

              'Calculate array index where current item data will be stored
12            lngArrayIndex = (AWS_ITEMSPERPAGE * (lngCurrentPage - 1) + lngCurrentItem)

              'ASIN
13            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:ASIN")
14            If Not objNode Is Nothing Then
15                udtAWSItems(lngArrayIndex).strASIN = objNode.Text
16            End If

              'Small image URL
17            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:SmallImage/aws:URL")
18            If Not objNode Is Nothing Then
19                udtAWSItems(lngArrayIndex).strSmallImageURL = objNode.Text
20                blnSmallImageAvailable = True
21            End If

              'Small image height
22            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:SmallImage/aws:Height")
23            If Not objNode Is Nothing Then
24                udtAWSItems(lngArrayIndex).lngSmallImageHeight = objNode.Text
25            End If

              'Small image width
26            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:SmallImage/aws:Width")
27            If Not objNode Is Nothing Then
28                udtAWSItems(lngArrayIndex).lngSmallImageWidth = objNode.Text
29            End If

              'Medium image URL
30            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:MediumImage/aws:URL")
31            If Not objNode Is Nothing Then
32                udtAWSItems(lngArrayIndex).strMediumImageURL = objNode.Text
33            End If

              'Medium image height
34            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:MediumImage/aws:Height")
35            If Not objNode Is Nothing Then
36                udtAWSItems(lngArrayIndex).lngMediumImageHeight = objNode.Text
37            End If

              'Medium image width
38            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:MediumImage/aws:Width")
39            If Not objNode Is Nothing Then
40                udtAWSItems(lngArrayIndex).lngMediumImageWidth = objNode.Text
41            End If

              'Large image URL
42            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:LargeImage/aws:URL")
43            If Not objNode Is Nothing Then
44                udtAWSItems(lngArrayIndex).strLargeImageURL = objNode.Text
45            End If

              'Large image height
46            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:LargeImage/aws:Height")
47            If Not objNode Is Nothing Then
48                udtAWSItems(lngArrayIndex).lngLargeImageHeight = objNode.Text
49            End If

              'Large image width
50            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:LargeImage/aws:Width")
51            If Not objNode Is Nothing Then
52                udtAWSItems(lngArrayIndex).lngLargeImageWidth = objNode.Text
53            End If

              'Artist
54            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:ItemAttributes/aws:Artist")
55            If Not objNode Is Nothing Then
56                udtAWSItems(lngArrayIndex).strArtist = objNode.Text
57            End If

              'Product group
58            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:ItemAttributes/aws:ProductGroup")
59            If Not objNode Is Nothing Then
60                udtAWSItems(lngArrayIndex).strProductGroup = objNode.Text
61            End If

              'Title
62            Set objNode = objNodes(lngCurrentItem).selectSingleNode("aws:ItemAttributes/aws:Title")
63            If Not objNode Is Nothing Then
64                udtAWSItems(lngArrayIndex).strTitle = objNode.Text
65            End If

              'If the cancel flag was not set increment the item counter, otherwise stop
66            If Not blnCancel Then
67                lngCurrentItem = lngCurrentItem + 1
68              Else
69                objXMLParser.abort
70                ParseResponse = ACD_USERSTOP
71                Exit Function
72            End If
73        Loop Until lngCurrentItem > lngTotalNumberOfItems

          'Return OK
74        ParseResponse = ACD_OK

75        On Error GoTo 0

76    Exit Function

'Handle errors
ErrorHandler:
77        udtRuntimeError.lngLine = Erl
78        udtRuntimeError.lngNumber = Err.Number
79        udtRuntimeError.strDescription = Err.Description
80        udtRuntimeError.strSource = Err.Source
81        If udtRuntimeError.strSource = "AmazonCoverDownloaderCore" Then
82            udtRuntimeError.strSource = udtRuntimeError.strSource & " :: clsMain :: ParseResponse()"
83        End If

84        ParseResponse = ACD_PARSERUNTIMEERROR

End Function

Public Function Query(ByVal strQueryString As String, Optional ByVal strSubscriptionID As String = "1V77S1FA258XCHV2VHR2", Optional ByVal strLocaleURL As String = "amazon.com", Optional ByVal strMediaType As String = "Music", Optional ByVal lngNumberOfPagesToProcess As Long = 0) As Long

        Dim lngCurrentPage As Long
        Dim lngExternalFunctionResult As Long
        Dim lngTotalNumberOfPages As Long
        Dim objNode As IXMLDOMNode
        Dim strURL As String

1         On Error GoTo ErrorHandler

          'Stop if the XML parser is not instantiated
2         If objXMLParser Is Nothing Then
3             Query = ACD_NOXMLPARSER
4             RaiseEvent Done(Query)
5             Exit Function
6         End If

          'Reset global variables
7         blnCancel = False
8         blnSmallImageAvailable = False
9         Erase udtAWSItems
10        ZeroMemory udtAWSError, LenB(udtAWSError)
11        ZeroMemory udtRuntimeError, LenB(udtRuntimeError)

          'Set page counters to 1
12        lngCurrentPage = 1
13        lngTotalNumberOfPages = 1

          'Inform the caller that work started
14        RaiseEvent WorkStarted

          'Loop through all result pages
15        Do
              'Inform the caller about the progress
16            RaiseEvent Progress(lngCurrentPage, lngTotalNumberOfPages)

              'Check if there are queued message and parse them
17            If GetQueueStatus(QS_ALLINPUT) Then
18                DoEvents
19            End If

              'Perform REST query and stop if the document was not loaded
20            strURL = "http://webservices." & strLocaleURL & "/onca/xml?Service=AWSECommerceService&ItemPage=" & lngCurrentPage & "&SearchIndex=" & strMediaType & "&SubscriptionId=" & strSubscriptionID & "&Keywords=" & strQueryString & "&Version=2005-10-05&ResponseGroup=Small,Images&Operation=ItemSearch"
21            If Not objXMLParser.Load(strURL) Then
22                Query = ACD_NOLOADXML
23                RaiseEvent Done(Query)
24                Exit Function
25            End If

              'Wait until XML document was loaded
26            Do
                  'If the cancel flag was not set check if there are queued message and parse them, otherwise stop
27                If Not blnCancel Then
28                    If GetQueueStatus(QS_ALLINPUT) Then
29                        DoEvents
30                    End If
31                  Else
32                    objXMLParser.abort
33                    Query = ACD_USERSTOP
34                    RaiseEvent Done(Query)
35                    Exit Function
36                End If
37            Loop Until blnXMLLoaded

              'Validate query and stop if validation failed
38            lngExternalFunctionResult = ValidateQuery
39            If Not lngExternalFunctionResult = ACD_OK Then
40                Query = lngExternalFunctionResult
41                RaiseEvent Done(Query)
42                Exit Function
43            End If

              'If we're at the first page, retrieve the number of items and pages
44            If lngCurrentPage = 1 Then

                  'Obtain number of items and stop on error (the number of items is either calculated based on the number of pages to parse or retrieved from AWS directly)
45                Set objNode = objXMLParser.selectSingleNode("aws:ItemSearchResponse/aws:Items/aws:TotalResults")
46                If Not objNode Is Nothing Then
47                    If lngNumberOfPagesToProcess = 0 Then
48                        ReDim udtAWSItems(objNode.Text - 1)
49                      Else
50                        ReDim udtAWSItems(AWS_ITEMSPERPAGE * Abs(lngNumberOfPagesToProcess) - 1)
51                    End If
52                  Else
53                    Query = ACD_NOITEMCOUNT
54                    RaiseEvent Done(Query)
55                    Exit Function
56                End If

                  'Obtain number of pages and stop on error (the number of pages is either calculated based on the number of pages to parse or retrieved from AWS directly)
57                Set objNode = objXMLParser.selectSingleNode("aws:ItemSearchResponse/aws:Items/aws:TotalPages")
58                If Not objNode Is Nothing Then
59                    If lngNumberOfPagesToProcess = 0 Then
60                        lngTotalNumberOfPages = objNode.Text
61                      Else
62                        If objNode.Text >= Abs(lngNumberOfPagesToProcess) Then
63                            lngTotalNumberOfPages = Abs(lngNumberOfPagesToProcess)
64                          Else
65                            lngTotalNumberOfPages = objNode.Text
66                        End If
67                    End If
68                  Else
69                    Query = ACD_NOPAGECOUNT
70                    RaiseEvent Done(Query)
71                    Exit Function
72                End If

                  'Inform caller about the progress and that precise information about the total number of pages is available
73                RaiseEvent ProgressDataAvailable
74                RaiseEvent Progress(lngCurrentPage, lngTotalNumberOfPages)

                  'Check if there are queued message and parse them
75                If GetQueueStatus(QS_ALLINPUT) Then
76                    DoEvents
77                End If

78            End If

              'Parse page and stop on error
79            lngExternalFunctionResult = ParseResponse(lngCurrentPage)
80            If Not lngExternalFunctionResult = ACD_OK Then
81                Query = lngExternalFunctionResult
82                RaiseEvent Done(Query)
83                Exit Function
84            End If

              'If the cancel flag was not set increment the page counter, otherwise stop
85            If Not blnCancel Then
86                lngCurrentPage = lngCurrentPage + 1
87              Else
88                objXMLParser.abort
89                Query = ACD_USERSTOP
90                RaiseEvent Done(Query)
91                Exit Function
92            End If
93        Loop Until lngCurrentPage > lngTotalNumberOfPages

          'Check small image availability
94        If blnSmallImageAvailable Then
95            Query = ACD_OK
96            RaiseEvent Done(Query)
97          Else
98            Query = ACD_NOSMALLIMAGE
99            RaiseEvent Done(Query)
100       End If

101       On Error GoTo 0

102   Exit Function

'Handle errors
ErrorHandler:
103       udtRuntimeError.lngLine = Erl
104       udtRuntimeError.lngNumber = Err.Number
105       udtRuntimeError.strDescription = Err.Description
106       udtRuntimeError.strSource = Err.Source
107       If udtRuntimeError.strSource = "AmazonCoverDownloaderCore" Then
108           udtRuntimeError.strSource = udtRuntimeError.strSource & " :: clsMain :: Query()"
109       End If

110       Query = ACD_QUERYRUNTIMEERROR
111       RaiseEvent Done(Query)

End Function

Public Property Get RuntimeError() As RuntimeError

1         RuntimeError = udtRuntimeError

End Property

Public Property Get StatusCodeDescription(ByVal lngStatusCode As Long) As String

1         Select Case lngStatusCode
            Case ACD_OK
2             StatusCodeDescription = "OK."
3           Case ACD_NOXMLPARSER
4             StatusCodeDescription = "The XML parser was not instantiated correctly."
5           Case ACD_NOLOADXML
6             StatusCodeDescription = "The XML parser was unable to load the requested document."
7           Case ACD_AWSERROR
8             StatusCodeDescription = "Amazon Web Services returned an error."
9           Case ACD_NOVALIDITYCHECK
10            StatusCodeDescription = "Unable to check the validity of the Amazon Web Services request."
11          Case ACD_NOITEMCOUNT
12            StatusCodeDescription = "Unable to retrieve the total number of results."
13          Case ACD_NORETRIEVEITEM
14            StatusCodeDescription = "Unable to retrieve the item list."
15          Case ACD_NOSMALLIMAGE
16            StatusCodeDescription = "Matches were found, but none of them contained an image."
17          Case ACD_NOPAGECOUNT
18            StatusCodeDescription = "Unable to retrieve the total number of pages."
19          Case ACD_QUERYRUNTIMEERROR
20            StatusCodeDescription = "A runtime error occurred while attempting to query Amazon Web Services."
21          Case ACD_PARSERUNTIMEERROR
22            StatusCodeDescription = "A runtime error occurred while attempting to parse the Amazon Web Services response."
23          Case ACD_VALIDATERUNTIMEERROR
24            StatusCodeDescription = "A runtime error occurred while attempting to validate the Amazon Web Services query."
25          Case ACD_USERSTOP
26            StatusCodeDescription = "User initialized stop."
27          Case Else
28            StatusCodeDescription = "Unknown error."
29        End Select

End Property

Private Function ValidateQuery() As Long

        Dim objNode As IXMLDOMNode

1         On Error GoTo ErrorHandler

          'Check if the query is valid
2         Set objNode = objXMLParser.selectSingleNode("aws:ItemSearchResponse/aws:Items/aws:Request/aws:IsValid")
3         If objNode Is Nothing Then
4             ValidateQuery = ACD_NOVALIDITYCHECK
5             Exit Function
6         End If

          'Check if there are errors
7         Set objNode = objXMLParser.selectSingleNode("aws:ItemSearchResponse/aws:Items/aws:Request/aws:Errors/aws:Error")
8         If Not objNode Is Nothing Then
              'Error code
9             Set objNode = objXMLParser.selectSingleNode("aws:ItemSearchResponse/aws:Items/aws:Request/aws:Errors/aws:Error/aws:Code")
10            If Not objNode Is Nothing Then
11                udtAWSError.strCode = objNode.Text
12            End If

              'Error description
13            Set objNode = objXMLParser.selectSingleNode("aws:ItemSearchResponse/aws:Items/aws:Request/aws:Errors/aws:Error/aws:Message")
14            If Not objNode Is Nothing Then
15                udtAWSError.strMessage = objNode.Text
16            End If

              'Return error code
17            ValidateQuery = ACD_AWSERROR
18            Exit Function
19        End If

          'Check if there are queued message and parse them
20        If GetQueueStatus(QS_ALLINPUT) Then
21            DoEvents
22        End If

          'Return OK
23        ValidateQuery = ACD_OK

24        On Error GoTo 0

25    Exit Function

'Handle errors
ErrorHandler:
26        udtRuntimeError.lngLine = Erl
27        udtRuntimeError.lngNumber = Err.Number
28        udtRuntimeError.strDescription = Err.Description
29        udtRuntimeError.strSource = Err.Source
30        If udtRuntimeError.strSource = "AmazonCoverDownloaderCore" Then
31            udtRuntimeError.strSource = udtRuntimeError.strSource & " :: clsMain :: ValidateQuery()"
32        End If

33        ValidateQuery = ACD_VALIDATERUNTIMEERROR

End Function

Public Property Get XMLError() As IXMLDOMParseError

1         If Not objXMLParser Is Nothing Then
2             Set XMLError = objXMLParser.parseError
3         End If

End Property
