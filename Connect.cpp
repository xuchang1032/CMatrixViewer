// Connect.cpp : Implementation of CConnect
#include "stdafx.h"
#include "AddIn.h"
#include "Connect.h"

extern CAddInModule _AtlModule;

// CConnect
STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, ext_ConnectMode ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** /*custom*/ )
{
	HRESULT hr = S_OK;
	pApplication->QueryInterface(__uuidof(DTE2), (LPVOID*)&m_pDTE);
	pAddInInst->QueryInterface(__uuidof(AddIn), (LPVOID*)&m_pAddInInstance);
	if(ConnectMode == 5) //5 == ext_cm_UISetup
	{
		HRESULT hr = S_OK;
		CComPtr<IDispatch> pDisp;
		CComQIPtr<Commands> pCommands;
		CComQIPtr<Commands2> pCommands2;
		CComQIPtr<_CommandBars> pCommandBars;
		CComPtr<CommandBarControl> pCommandBarControl;
		CComPtr<Command> pCreatedCommand;
		CComPtr<CommandBar> pMenuBarCommandBar;
		CComPtr<CommandBarControls> pMenuBarControls;
		CComPtr<CommandBarControl> pToolsCommandBarControl;
		CComQIPtr<CommandBar> pToolsCommandBar;
		CComQIPtr<CommandBarPopup> pToolsPopup;

		IfFailGoCheck(m_pDTE->get_Commands(&pCommands), pCommands);
		pCommands2 = pCommands;
		if(SUCCEEDED(pCommands2->AddNamedCommand2(m_pAddInInstance, CComBSTR("CMatrixViewer"), CComBSTR("View CMatrix"), CComBSTR("View CMatrix<T> In Output Window"), VARIANT_TRUE, CComVariant(59), NULL, vsCommandStatusSupported+vsCommandStatusEnabled, vsCommandStylePictAndText, vsCommandControlTypeButton, &pCreatedCommand)) && (pCreatedCommand))
		{
			//Add a button to the tools menu bar.
			IfFailGoCheck(m_pDTE->get_CommandBars(&pDisp), pDisp);
			pCommandBars = pDisp;

			//Find the MenuBar command bar, which is the top-level command bar holding all the main menu items:
			IfFailGoCheck(pCommandBars->get_Item(CComVariant(L"Code Window"), &pMenuBarCommandBar), pMenuBarCommandBar);
// 			IfFailGoCheck(pMenuBarCommandBar->get_Controls(&pMenuBarControls), pMenuBarControls);
// 
// 			//Find the Tools command bar on the MenuBar command bar:
// 			IfFailGoCheck(pMenuBarControls->get_Item(CComVariant(L"Tools"), &pToolsCommandBarControl), pToolsCommandBarControl);
// 			pToolsPopup = pToolsCommandBarControl;
// 			IfFailGoCheck(pToolsPopup->get_CommandBar(&pToolsCommandBar), pToolsCommandBar);

			//Add the control:
			pDisp = NULL;
			IfFailGoCheck(pCreatedCommand->AddControl(pMenuBarCommandBar, 5, &pDisp), pDisp);
		}
		return S_OK;
	}
Error:
	return hr;
}

STDMETHODIMP CConnect::OnDisconnection(ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
	m_pDTE = NULL;
	m_pAddInInstance = NULL;
	return S_OK;
}

STDMETHODIMP CConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CConnect::QueryStatus(BSTR bstrCmdName, vsCommandStatusTextWanted NeededText, vsCommandStatus *pStatusOption, VARIANT *pvarCommandText)
{
  if(NeededText == vsCommandStatusTextWantedNone)
	{
	  if(!_wcsicmp(bstrCmdName, L"CMatrixViewer.Connect.CMatrixViewer"))
	  {
		  *pStatusOption = (vsCommandStatus)(vsCommandStatusEnabled+vsCommandStatusSupported);
	  }
  }
	return S_OK;
}

STDMETHODIMP CConnect::Exec(BSTR bstrCmdName, vsCommandExecOption ExecuteOption, VARIANT * /*pvarVariantIn*/, VARIANT * /*pvarVariantOut*/, VARIANT_BOOL *pvbHandled)
{
	*pvbHandled = VARIANT_FALSE;
	if(ExecuteOption == vsCommandExecOptionDoDefault)
	{
		if(!_wcsicmp(bstrCmdName, L"CMatrixViewer.Connect.CMatrixViewer"))
		{
			ViewCMatrix();
			*pvbHandled = VARIANT_TRUE;
			return S_OK;
		}
	}
	return S_OK;
}

HRESULT CConnect::ViewCMatrix()
{
	HRESULT hr = S_OK;

	//获取OutputWindow
	ToolWindows* pToolWins = NULL;
	IfFailGoCheck(m_pDTE->get_ToolWindows(&pToolWins), pToolWins);
	OutputWindow* pOutputWindow = NULL;
	IfFailGoCheck(pToolWins->get_OutputWindow(&pOutputWindow), pOutputWindow);
	OutputWindowPanes* pOutputWindowPanes = NULL;
	IfFailGoCheck(pOutputWindow->get_OutputWindowPanes(&pOutputWindowPanes), pOutputWindowPanes);
	OutputWindowPane* pDebugPane = NULL;
	IfFailGoCheck(pOutputWindowPanes->Item(CComVariant("Debug"), &pDebugPane), pDebugPane);

	//获取选中文本
	Document* pActDocument = NULL;
	IfFailGoCheck(m_pDTE->get_ActiveDocument(&pActDocument), pActDocument);
	TextSelection* pTextSelection = NULL;
	IfFailGoCheck(pActDocument->get_Selection((IDispatch**)&pTextSelection), pTextSelection);
	VARIANT_BOOL hasText;
	pTextSelection->get_IsEmpty(&hasText);
	if (hasText == VARIANT_TRUE)	return hr;//无选中文字则返回
	BSTR bSelText;
	pTextSelection->get_Text(&bSelText);
	char cMatrixName[256] = {0};//选中的字符串
	char* lpszText = _com_util::ConvertBSTRToString(bSelText);
	::OutputDebugString(lpszText);
	strcpy_s(cMatrixName, lpszText);
	delete []lpszText;

	//检查分本是否是正确的CMatrix或CVector
	BSTR bExpression, bExpName, bExpValue, bExpType;
	char cExpression[256] = {0};
	Debugger* pDebugger = NULL;
	IfFailGoCheck(m_pDTE->get_Debugger(&pDebugger), pDebugger);

	//检查选中字符串,cMatrixName应该是一个可用的Expression
	Expression* pExpression = NULL;
	strcpy_s(cExpression, cMatrixName);
	bExpression = _com_util::ConvertStringToBSTR(cExpression);
	IfFailGoCheck(pDebugger->GetExpression(bExpression, VARIANT_FALSE, 100, &pExpression), pExpression);
	::SysFreeString(bExpression);

	//区分CMatrix<i4>, CMatrix<f8>,获取行数，列数
	int nRowNumber = 0, nColNumber = 0;
	BOOL isDouble = FALSE;
	pExpression->get_Name(&bExpName);
	pExpression->get_Type(&bExpType);

	char cPartExpression[100] = {0};//(****).m_i4RowNumber,括号内的部分表达式
	if (wcscmp(bExpType, L"CMatrix<double>") == 0 ||
		wcscmp(bExpType, L"CMatrix<double> &") == 0 ||
		wcscmp(bExpType, L"const CMatrix<double>") == 0  || 
		wcscmp(bExpType, L"const CMatrix<double> &") == 0)			//CMatrix<f8>或者CMatrix<double> &
	{
		//形如：(maF8).m_i4RowNumber
		sprintf(cPartExpression, "%s", cMatrixName);
		isDouble = TRUE;
	}
	else if(wcscmp(bExpType, L"CMatrix<double> *") == 0 ||
		wcscmp(bExpType, L"const CMatrix<double> *") == 0)	//CMatrix<f8> *
	{
		//(*(pmaF8)).m_i4RowNumber
		sprintf(cPartExpression, "*(%s)", cMatrixName);
		isDouble = TRUE;
	}
	else if(wcscmp(bExpType, L"CMatrix<int>") == 0 || 
		wcscmp(bExpType, L"CMatrix<int> &") == 0 || 
		wcscmp(bExpType, L"const CMatrix<int>") == 0 ||
		wcscmp(bExpType, L"const CMatrix<int> &") == 0)				//CMatrix<i4>
	{
		//(maI4).m_i4RowNumber
		sprintf(cPartExpression, "%s", cMatrixName);
		isDouble = FALSE;
	}
	else if(wcscmp(bExpType, L"CMatrix<int> *") == 0 ||
		wcscmp(bExpType, L"const CMatrix<int> *") == 0)			//CMatrix<i4> *
	{
		//(*(pmaI4)).m_i4RowNumber
		sprintf(cPartExpression, "*(%s)", cMatrixName);
		isDouble = FALSE;
	}
	else if(wcscmp(bExpType, L"CVector<double>") == 0 ||
		wcscmp(bExpType, L"CVector<double> &") == 0 ||
		wcscmp(bExpType, L"const CVector<double>") == 0  || 
		wcscmp(bExpType, L"const CVector<double> &") == 0)		//CVector<f8>
	{
		//((*(CMatrix<double>*)(&veF8))).m_i4RowNumber
		sprintf(cPartExpression, "(*(CMatrix<double>*)(&%s))", cMatrixName);
		isDouble = TRUE;
	}
	else if(wcscmp(bExpType, L"CVector<double> *") == 0 ||
		wcscmp(bExpType, L"const CVector<double> *") == 0)	//CVector<f8> *
	{
		//((*(CMatrix<double>*)(&*(pveF8)))).m_i4RowNumber
		sprintf(cPartExpression, "(*(CMatrix<double>*)(&*(%s)))", cMatrixName);
		isDouble = TRUE;
	}
	else if(wcscmp(bExpType, L"CVector<int>") == 0 || 
		wcscmp(bExpType, L"CVector<int> &") == 0 || 
		wcscmp(bExpType, L"const CVector<int>") == 0 ||
		wcscmp(bExpType, L"const CVector<int> &") == 0)				//CVector<i4>
	{
		//((*(CMatrix<int>*)(&veI4))).m_i4RowNumber
		sprintf(cPartExpression, "(*(CMatrix<int>*)(&%s))", cMatrixName);
		isDouble = FALSE;
	}
	else if(wcscmp(bExpType, L"CVector<int> *") == 0 ||
		wcscmp(bExpType, L"const CVector<int> *") == 0)			//CVector<i4> *
	{
		//((*(CMatrix<int>*)(&*(pveI4)))).m_i4RowNumber
		sprintf(cPartExpression, "(*(CMatrix<int>*)(&*(%s)))", cMatrixName);
		isDouble = FALSE;
	}
	::SysFreeString(bExpName);
	::SysFreeString(bExpType);

	sprintf(cExpression, "(%s).m_i4RowNumber", cPartExpression);
	bExpression = _com_util::ConvertStringToBSTR(cExpression);
	IfFailGoCheck(pDebugger->GetExpression(bExpression, VARIANT_FALSE, 100, &pExpression), pExpression);
	::SysFreeString(bExpression);
	pExpression->get_Value(&bExpValue);
	nRowNumber = _wtoi(bExpValue);
	::SysFreeString(bExpValue);

	sprintf(cExpression, "(%s).m_i4ColumnNumber", cPartExpression);
	bExpression = _com_util::ConvertStringToBSTR(cExpression);
	IfFailGoCheck(pDebugger->GetExpression(bExpression, VARIANT_FALSE, 100, &pExpression), pExpression);
	::SysFreeString(bExpression);
	pExpression->get_Value(&bExpValue);
	nColNumber = _wtoi(bExpValue);
	::SysFreeString(bExpValue);

	//打印行列
	pDebugPane->Activate();
	::OutputDebugString("\n");
	char szOutput[4096] = {0};
	strcat_s(szOutput, "\n");
	strcat_s(szOutput, cMatrixName);
	char szRow[10] = {0}, szCol[10] = {0};
	sprintf(szRow, "%d", nRowNumber);
	sprintf(szCol, "%d", nColNumber);
	strcat_s(szOutput, " ( ");
	strcat_s(szOutput, szRow);
	strcat_s(szOutput, " x ");
	strcat_s(szOutput, szCol);
	strcat_s(szOutput, " ) ");
	strcat_s(szOutput, "\n");
	int t=nColNumber; while(t!=0) { strcat_s(szOutput, "-----------"); t--;}
	strcat_s(szOutput, "\n");
	BSTR bOutput = _com_util::ConvertStringToBSTR(szOutput);
	pDebugPane->OutputString(bOutput);
	::SysFreeString(bOutput);
	memset(szOutput, 0, 4096);
	for (int r=0; r<nRowNumber; r++)
	{
		for (int c=0; c<nColNumber; c++)
		{
			//形如：(****).m_pData[n],****为以上cPartExpression
			char szExp[50] = {0};
			sprintf(szExp, "(%s).m_pData[%d*%d+%d]", cPartExpression, r, nColNumber, c);
			bExpression = _com_util::ConvertStringToBSTR(szExp);
			IfFailGoCheck(pDebugger->GetExpression(bExpression, false, 100, &pExpression), pExpression);
			pExpression->get_Value(&bExpValue);

			char cShort[50] = {0};
			if (isDouble)
			{
				double dValue = 0;
				dValue = _wtof(bExpValue);
				sprintf(cShort, "%10.4f ", dValue);
			}
			else
			{
				int nValue = 0;
				nValue = _wtoi(bExpValue);
				sprintf(cShort, "%10d", nValue);
			}
			if(c+1 == nColNumber) strcat_s(cShort, "\n");
			strcat_s(szOutput, cShort);

			::SysFreeString(bExpression);
			::SysFreeString(bExpValue);
		}
		bOutput = _com_util::ConvertStringToBSTR(szOutput);
		pDebugPane->OutputString(bOutput);
		::SysFreeString(bOutput);
		memset(szOutput, 0, 4096);
	}
	t=nColNumber; while(t!=0) { strcat_s(szOutput, "-----------"); t--;}
	strcat_s(szOutput, "\n");

	::OutputDebugString(szOutput);
	bOutput = _com_util::ConvertStringToBSTR(szOutput);
	pDebugPane->OutputString(bOutput);
	::SysFreeString(bOutput);

Error:
	return hr;
}