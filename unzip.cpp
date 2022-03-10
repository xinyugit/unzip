#define _NO_CRT_STDIO_INLINE
#include <Windows.h>
#include <Shldisp.h> ///Shldisp.idl
#include <stdio.h>
#include <locale.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "User32.lib")
#pragma comment(lib, "kernel32.lib")
#pragma comment(lib, "msvcrt64.lib")
//#pragma comment(linker, "/ENTRY:main")

bool UnZipFolder(wchar_t *zipFile, wchar_t *destination)
{

   DWORD strlen = 0;
   HRESULT hResult;
   IShellDispatch *pISD;
   Folder *pToFolder = NULL;
   Folder *pFromFolder = NULL;
   FolderItems *pFolderItems = NULL;
   FolderItem *pItem = NULL;

   VARIANT vDir, vFile, vOpt;
   BSTR strptr1, strptr2;
   CoInitialize(NULL);

   bool bReturn = false;

   hResult = CoCreateInstance(CLSID_Shell, NULL, CLSCTX_INPROC_SERVER, IID_IShellDispatch, (void **)&pISD);

   if (FAILED(hResult))
   {
      return bReturn;
   }

   VariantInit(&vOpt);
   vOpt.vt = VT_I4;
   vOpt.lVal = 16 + 4; // Do not display a progress dialog box ~ This will not work properly!

   strptr1 = SysAllocString(zipFile);
   strptr2 = SysAllocString(destination);

   VariantInit(&vFile);
   vFile.vt = VT_BSTR;
   vFile.bstrVal = strptr1;
   hResult = pISD->NameSpace(vFile, &pFromFolder);

   VariantInit(&vDir);
   vDir.vt = VT_BSTR;
   vDir.bstrVal = strptr2;

   hResult = pISD->NameSpace(vDir, &pToFolder);

   if (S_OK == hResult)
   {
      hResult = pFromFolder->Items(&pFolderItems);
      if (SUCCEEDED(hResult))
      {
         long lCount = 0;
         pFolderItems->get_Count(&lCount);
         IDispatch *pDispatch = NULL;
         pFolderItems->QueryInterface(IID_IDispatch, (void **)&pDispatch);
         VARIANT vtDispatch;
         VariantInit(&vtDispatch);
         vtDispatch.vt = VT_DISPATCH;
         vtDispatch.pdispVal = pDispatch;

         // cout << "Extracting files ...\n";
         hResult = pToFolder->CopyHere(vtDispatch, vOpt);
         if (hResult != S_OK)
            return false;

         // Cross check and wait until all files are zipped!
         FolderItems *pToFolderItems;
         hResult = pToFolder->Items(&pToFolderItems);

         if (S_OK == hResult)
         {
            long lCount2 = 0;

            hResult = pToFolderItems->get_Count(&lCount2);
            if (S_OK != hResult)
            {
               pFolderItems->Release();
               pToFolderItems->Release();
               SysFreeString(strptr1);
               SysFreeString(strptr2);
               pISD->Release();
               CoUninitialize();
               return false;
            }
            // Use this code in a loop if you want to cross-check the items unzipped.
            /*if(lCount2 != lCount)
            {
             pFolderItems->Release();
             pToFolderItems->Release();
             SysFreeString(strptr1);
             SysFreeString(strptr2);
             pISD->Release();
             CoUninitialize();
             return false;
            }*/

            bReturn = true;
         }

         pFolderItems->Release();
         pToFolderItems->Release();
      }

      pToFolder->Release();
      pFromFolder->Release();
   }

   // cout << "Over!\n";
   SysFreeString(strptr1);
   SysFreeString(strptr2);
   pISD->Release();

   CoUninitialize();
   return bReturn;
}

int wmain(int argc, wchar_t *argv[])
{

   setlocale(LC_ALL, "");

   if (argc != 3)
   {
      printf("\r\n 最简单的ZIP解压程序 \r\n");
      printf("\r\n unzip zip文件 解压的目录\r\n");
      return 0;
   }

   // UnZipFolder(L"F:\\zipfldr\\test.zip", L"F:\\zipfldr\\test");

   wchar_t zip[260] = {0};
   wchar_t dir[260] = {0};

   GetCurrentDirectoryW(260, zip);
   GetCurrentDirectoryW(260, dir);

   if (wcsstr(argv[1], L":") == 0) //带盘符的路径
   {
      wcscat(zip, L"\\");
      wcscat(zip, argv[1]);
   }
   else
   {
      wcscpy(zip, argv[1]);
   }

   if (wcsstr(argv[2], L":") == 0) //带盘符的路径
   {
      wcscat(dir, L"\\");
      wcscat(dir, argv[2]);
   }
   else
   {
      wcscpy(dir, argv[2]);
   }

   CreateDirectoryW(dir, 0);

   wprintf(L"\r\n正在解压到 %s\r\n", dir);
   DWORD64 tk = GetTickCount64();
   UnZipFolder(zip, dir);
   wprintf(L"\r\n解压完成耗时: %llums\r\n", GetTickCount64() - tk);
   return 0;
}