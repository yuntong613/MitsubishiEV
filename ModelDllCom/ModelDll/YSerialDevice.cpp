#include "StdAfx.h"
#include "common.h"
#include "YSerialDevice.h"
#include "IniFile.h"
#include "ItemBrowseDlg.h"
#include "YSerialItem.h"
#include <cstringt.h>
#include "ModelDll.h"
#include "OPCIniFile.h"
#include "ModbusCRC.h"

extern CModelDllApp theApp;

CCriticalSection m_ComSec;

DWORD CALLBACK QuertyThread(LPVOID pParam)
{
	YSerialDevice* pDevice = (YSerialDevice*)pParam;
	while (!pDevice->m_bStop)
	{
		Sleep(1000);
		pDevice->QueryOnce();
	}
	return 0;
}


YSerialDevice::YSerialDevice(LPCSTR pszAppPath)
	: m_nBaudRate(0)
	, m_bStop(true)
{
	m_nParity = 0;
	m_hQueryThread = INVALID_HANDLE_VALUE;
	y_lUpdateTimer = 0;
	m_nUseLog = 0;
	CString strConfigFile(pszAppPath);
	strConfigFile += _T("\\ComFile.ini");
	if (!InitConfig(strConfigFile))
	{
		return;
	}

	// 	CString strListItemsFile(pszAppPath);
	// 	strListItemsFile += _T("\\ListItems.ini");
	// 	COPCIniFile opcFile;
	// 	if (!opcFile.Open(strListItemsFile,CFile::modeRead|CFile::typeText))
	// 	{
	// 		AfxMessageBox("Can't open INI file!");
	// 		return;
	// 	}
	// 	CArchive ar(&opcFile,CArchive::load);
	// 	Serialize(ar);
	// 	opcFile.Close();
}

YSerialDevice::~YSerialDevice(void)
{
	m_Com.Close();
	POSITION pos = m_ItemsArray.GetStartPosition();
	YSerialItem* pItem = NULL;
	CString strItemName;
	while (pos) {
		m_ItemsArray.GetNextAssoc(pos, strItemName, (CObject*&)pItem);
		if (pItem)
		{
			delete pItem;
			pItem = NULL;
		}
	}
	m_ItemsArray.RemoveAll();
}

void YSerialDevice::Serialize(CArchive& ar)
{
	if (ar.IsStoring()) {
	}
	else {
		Load(ar);
	}
}

BOOL YSerialDevice::InitConfig(CString strFilePath)
{
	if (!PathFileExists(strFilePath))
		return FALSE;
	m_strConfigFile = strFilePath;

	CIniFile iniFile(m_strConfigFile);
	m_lRate = iniFile.GetUInt("param", "UpdateRate", 3000);
	m_nUseLog = iniFile.GetUInt("param", "Log", 0);

	m_nBaudRate = iniFile.GetUInt("ComInfo", "BaudRate", 9600);
	m_nParity = iniFile.GetUInt("ComInfo", "Parity", 0);

	MakeItems();
	return TRUE;
}

void YSerialDevice::Load(CArchive& ar)
{
	/*	LoadItems(ar);*/


}
void YSerialDevice::MakeItems()
{
	YOPCItem* pItemStatus = NULL;
	DWORD dwItemPId = 0L;
	CString strSec, strItemName, strItemDesc;

	CString strComPort, strAddr;
	int bAddr = 0;

	CIniFile iniFile(m_strConfigFile);
	CStringArray ComPortArray, AddrArray;
	iniFile.GetArray("ComInfo", "ComPort", &ComPortArray);

	for (int i = 0; i < ComPortArray.GetCount(); i++)
	{
		AddrArray.RemoveAll();
		strComPort = ComPortArray.GetAt(i);
		strSec.Format("COM%d", atoi(strComPort));

		iniFile.GetArray(strSec, "Addr", &AddrArray);

		for (int k = 0; k < AddrArray.GetCount(); k++)
		{
			StrToIntEx("0x" + AddrArray.GetAt(k), STIF_SUPPORT_HEX, &bAddr);

			//运行楼层
			strItemName.Format("%s_FR_%02X_S", strSec, bAddr);
			strItemDesc.Format("%d站运行楼层", bAddr);
			pItemStatus = new YStringItem(dwItemPId++, strItemName, strItemDesc);
			m_ItemsArray.SetAt(pItemStatus->GetName(), (CObject*)pItemStatus);

			//运行方向
			strItemName.Format("%s_UPDOWN_%02X_S",strSec, bAddr);
			strItemDesc.Format("%d站运行方向", bAddr);
			pItemStatus = new YShortItem(dwItemPId++, strItemName, strItemDesc);
			m_ItemsArray.SetAt(pItemStatus->GetName(), (CObject*)pItemStatus);

			//与电梯通讯状态
			strItemName.Format("%s_EV_%02X_S", strSec, bAddr);
			strItemDesc.Format("%d站电梯通讯", bAddr);
			pItemStatus = new YShortItem(dwItemPId++, strItemName, strItemDesc);
			m_ItemsArray.SetAt(pItemStatus->GetName(), (CObject*)pItemStatus);

			//与PC通讯状态
			strItemName.Format("%s_PC_%02X_S", strSec, bAddr);
			strItemDesc.Format("%d站PC通讯", bAddr);
			pItemStatus = new YShortItem(dwItemPId++, strItemName, strItemDesc);
			m_ItemsArray.SetAt(pItemStatus->GetName(), (CObject*)pItemStatus);

			//故障状态
			strItemName.Format("%s_GZ_%02X_S", strSec, bAddr);
			strItemDesc.Format("%d站故障状态", bAddr);
			pItemStatus = new YShortItem(dwItemPId++, strItemName, strItemDesc);
			m_ItemsArray.SetAt(pItemStatus->GetName(), (CObject*)pItemStatus);

			//维保状态
			strItemName.Format("%s_WB_%02X_S", strSec, bAddr);
			strItemDesc.Format("%d站维保状态", bAddr);
			pItemStatus = new YShortItem(dwItemPId++, strItemName, strItemDesc);
			m_ItemsArray.SetAt(pItemStatus->GetName(), (CObject*)pItemStatus);
		}
	}
}

void YSerialDevice::LoadItems(CArchive& ar)
{
	COPCIniFile* pIniFile = static_cast<COPCIniFile*>(ar.GetFile());
	YOPCItem* pItem = NULL;
	int nItems = 0;
	CString strTmp("Item");
	CString strItemName;
	CString strItemDesc;
	CString strValue;
	DWORD dwItemPId = 0L;
	strTmp += CString(_T("List"));
	if (pIniFile->ReadNoSeqSection(strTmp)) {
		nItems = pIniFile->GetItemsCount(strTmp, "Item");
		for (int i = 0; i < nItems && !pIniFile->Endof(); i++)
		{
			try {
				if (pIniFile->ReadIniItem("Item", strTmp))
				{
					if (!pIniFile->ExtractSubValue(strTmp, strValue, 1))
						throw new CItemException(CItemException::invalidId, pIniFile->GetFileName());
					dwItemPId = atoi(strValue);
					if (!pIniFile->ExtractSubValue(strTmp, strItemName, 2))strItemName = _T("Unknown");
					if (!pIniFile->ExtractSubValue(strTmp, strItemDesc, 3)) strItemDesc = _T("Unknown");
					pItem = new YShortItem(dwItemPId, strItemName, strItemDesc);
					if (GetItemByName(strItemName))
						delete pItem;
					else
						m_ItemsArray.SetAt(pItem->GetName(), (CObject*)pItem);
				}
			}
			catch (CItemException* e) {
				if (pItem) delete pItem;
				e->Delete();
			}
		}
	}
}

void YSerialDevice::OnUpdate()
{

}

CString YSerialDevice::GetFloorName(BYTE bStatus)
{
	CString strFloorName;
	if (bStatus >= 1 && bStatus <= 48)
	{
		strFloorName.Format("%d", bStatus);
	}
	else if (bStatus >= 49 && bStatus <= 54)
	{
		strFloorName.Format("%d", 49 - (int)bStatus);
	}
	else if (bStatus >= 55 && bStatus <= 59)
	{
		strFloorName.Format("%dB", bStatus - 55 + 1);
	}
	else if (bStatus == 60)
	{
		strFloorName = "A";
	}
	else if (bStatus == 61)
	{
		strFloorName = "B";
	}
	else if (bStatus >= 62 && bStatus <= 67)
	{
		strFloorName.Format("B%d", bStatus - 62 + 1);
	}
	else if (bStatus == 68)
	{
		strFloorName = "C";
	}
	else if (bStatus == 69)
	{
		strFloorName = "D";
	}
	else if (bStatus == 70)
	{
		strFloorName = "E";
	}
	else if (bStatus == 71)
	{
		strFloorName = "G";
	}
	else if (bStatus >= 72 && bStatus <= 74)
	{
		strFloorName.Format("G%d", bStatus - 72 + 1);
	}
	else if (bStatus == 75)
	{
		strFloorName = "GF";
	}
	else if (bStatus == 76)
	{
		strFloorName = "H";
	}
	else if (bStatus == 77)
	{
		strFloorName = "K";
	}
	else if (bStatus == 78)
	{
		strFloorName = "L";
	}
	else if (bStatus >= 79 && bStatus <= 81)
	{
		strFloorName.Format("L%d", bStatus - 79 + 1);
	}
	else if (bStatus == 82)
	{
		strFloorName = "LB";
	}
	else if (bStatus == 83)
	{
		strFloorName = "LG";
	}
	else if (bStatus == 84)
	{
		strFloorName = "M";
	}
	else if (bStatus >= 85 && bStatus <= 86)
	{
		strFloorName.Format("M%d", bStatus - 85 + 1);
	}
	else if (bStatus == 90)
	{
		strFloorName = "M6";
	}
	else if (bStatus == 91)
	{
		strFloorName = "MB";
	}
	else if (bStatus == 92)
	{
		strFloorName = "P";
	}
	else if (bStatus >= 93 && bStatus <= 98)
	{
		strFloorName.Format("P%d", bStatus - 93);
	}
	else if (bStatus == 99)
	{
		strFloorName = "PB";
	}
	else if (bStatus == 100)
	{
		strFloorName = "PH";
	}
	else if (bStatus == 101)
	{
		strFloorName = "PL";
	}
	else if (bStatus == 102)
	{
		strFloorName = "PP";
	}
	else if (bStatus == 103)
	{
		strFloorName = "R";
	}
	else if (bStatus >= 104 && bStatus <= 106)
	{
		strFloorName.Format("R%d", bStatus - 104 + 1);
	}
	else if (bStatus == 107)
	{
		strFloorName = "S";
	}
	else if (bStatus >= 108 && bStatus <= 112)
	{
		strFloorName.Format("S%d", bStatus - 108 + 1);
	}
	else if (bStatus == 113)
	{
		strFloorName = "T";
	}
	else if (bStatus == 114)
	{
		strFloorName = "UB";
	}
	else if (bStatus == 115)
	{
		strFloorName = "UG";
	}

	return strFloorName;
}

int YSerialDevice::QueryOnce()
{
	y_lUpdateTimer += 1000;
	if (y_lUpdateTimer < m_lRate)
		return 0;
	y_lUpdateTimer = 0;

	CString strComPort, strSec;
	CIniFile iniFile(m_strConfigFile);
	CStringArray ComPortArray, AddrArray;
	iniFile.GetArray("ComInfo", "ComPort", &ComPortArray);

	for (int i = 0;i< ComPortArray.GetCount();i++)
	{
		AddrArray.RemoveAll();
		strComPort = ComPortArray.GetAt(i);
		strSec.Format("COM%d", atoi(strComPort));

		iniFile.GetArray(strSec, "Addr", &AddrArray);

		if (m_Com.Open(atoi(strComPort), m_nBaudRate, m_nParity))
		{
			CIniFile iniFile(m_strConfigFile);
			int nTimeOut = iniFile.GetInt("ComInfo", "TimeOut", 5000);
			int bAddr = 0;
			for (int k = 0; k < AddrArray.GetCount(); k++)
			{
				StrToIntEx("0x" + AddrArray.GetAt(k), STIF_SUPPORT_HEX, &bAddr);

				BYTE bSend[4] = { 0 };
				bSend[0] = 0xFA;
				bSend[1] = 0xFF;
				bSend[2] = bAddr;
				bSend[3] = 0XFE;

				m_ComSec.Lock();
				m_Com.Write(bSend, 4);

				CString strHex = Bin2HexStr(bSend, 4);
				OutPutLog("Send：" + strHex);

				BYTE recvBuf[10] = { 0 };
				DWORD dwRead = m_Com.Read(recvBuf, 6, nTimeOut);
				if (dwRead > 0)
				{
					strHex = Bin2HexStr(recvBuf, dwRead);
					OutPutLog("Recv：" + strHex);
				}

				m_ComSec.Unlock();

				if (dwRead == 6)
				{
					if (0xFA != recvBuf[0])
						return -1;
					if (0xFE != recvBuf[5])
						return -1;
					if (bAddr != recvBuf[1])
						return -1;
					if (0xFF != recvBuf[2])
						return -1;

					BYTE bFloorStatus = recvBuf[3];
					BYTE bOtherStatus = recvBuf[4];

					CString strItemName, strValue;
					YOPCItem* pItem = NULL;

					//运行楼层
					strItemName.Format("%s_FR_%02X_S", strSec, bAddr);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(GetFloorName(bFloorStatus));

					//运行方向 0为
					strItemName.Format("%s_UPDOWN_%02X_S",strSec, bAddr);
					int nUpDown = bOtherStatus & 3;
					if (nUpDown == 1)//上行
					{
						strValue = "1";
					}
					else if (nUpDown == 2) //	下行
					{
						strValue = "-1";
					}
					else if (nUpDown == 0) //停止
					{
						strValue = "0";
					}
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strValue);

					//与电梯通讯状态
					strItemName.Format("%s_EV_%02X_S", strSec, bAddr);
					strValue.Format("%d", (bOtherStatus >> 2) & 1);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strValue);

					//与PC通讯状态
					strItemName.Format("%s_PC_%02X_S", strSec, bAddr);
					strValue.Format("%d", (bOtherStatus >> 3) & 1);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strValue);

					//故障状态
					strItemName.Format("%s_GZ_%02X_S", strSec, bAddr);
					strValue.Format("%d", (bOtherStatus >> 4) & 1);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strValue);

					//维保状态
					strItemName.Format("%s_WB_%02X_S", strSec, bAddr);
					strValue.Format("%d", (bOtherStatus >> 5) & 1);
					pItem = GetItemByName(strItemName);
					if (pItem)
						pItem->OnUpdate(strValue);

				}
			}
			m_Com.Close();
		}
	}
	return 0;
}

void YSerialDevice::BeginUpdateThread()
{
	if (m_hQueryThread == INVALID_HANDLE_VALUE)
	{
		if (m_bStop == true)
		{
			m_bStop = false;
			m_hQueryThread = CreateThread(NULL, 0, QuertyThread, this, 0, NULL);
		}
	}
}

void YSerialDevice::EndUpdateThread()
{
	if (!m_bStop)
	{
		m_bStop = true;
		DWORD dwRet = WaitForSingleObject(m_hQueryThread, 3000);
		if (dwRet == WAIT_TIMEOUT)
			TerminateThread(m_hQueryThread, 0);

		m_hQueryThread = INVALID_HANDLE_VALUE;
	}
}

BYTE YSerialDevice::Hex2Bin(CString strHex)
{
	int iDec = 0;
	if (strHex.GetLength() == 2) {
		char cCh = strHex[0];
		if ((cCh >= '0') && (cCh <= '9'))iDec = cCh - '0';
		else if ((cCh >= 'A') && (cCh <= 'F'))iDec = cCh - 'A' + 10;
		else if ((cCh >= 'a') && (cCh <= 'f'))iDec = cCh - 'a' + 10;
		else return 0;
		iDec *= 16;
		cCh = strHex[1];
		if ((cCh >= '0') && (cCh <= '9'))iDec += cCh - '0';
		else if ((cCh >= 'A') && (cCh <= 'F'))iDec += cCh - 'A' + 10;
		else if ((cCh >= 'a') && (cCh <= 'f'))iDec += cCh - 'a' + 10;
		else return 0;
	}
	return iDec & 0xff;
}

int YSerialDevice::HexStr2Bin(BYTE * cpData, CString strData)
{
	CString strByte;
	for (int i = 0; i < strData.GetLength(); i += 2) {
		strByte = strData.Mid(i, 2);
		cpData[i / 2] = Hex2Bin(strByte);
	}
	return strData.GetLength() / 2;
}

CString YSerialDevice::Bin2HexStr(BYTE* cpData, int nLen)
{
	CString strResult, strTemp;
	for (int i = 0; i < nLen; i++)
	{
		strTemp.Format("%02X ", cpData[i]);
		strResult += strTemp;
	}
	return strResult;
}

void YSerialDevice::HandleData()
{

	return;
}

bool YSerialDevice::SetDeviceItemValue(CBaseItem* pAppItem)
{
	return false;
}

void YSerialDevice::OutPutLog(CString strMsg)
{
	if (m_nUseLog)
		m_Log.Write(strMsg);
}
