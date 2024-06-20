// Modified window event hook program, based on example from R. Chen
// Makes a text log entry whenever a task switch occurs, noting the new process of
// the window which gained focus, the presumed process of the presumed previous
// window which lost focus, and the time of the event.

// Standard C++ header files
//#include <stdio.h>
#include <iostream>
//#include <conio.h>

// Header file for Windows
// Enables strict typing in Windows.h
#define STRICT
#define NTDDL_VERSION 0x0A000007
#define _WIN32_WINNT 0x0A00
#define WINVER 0x0A00
// Should add LEAN_AND_MEAN and some NOxxx defines here later for optimization
// For now, just include everything and worry about optimizing later
#include <windows.h>
// Header file for multimedia system, needed for PlaySound if pruning windows.h
 #include <mmsystem.h>
//#include <combaseapi.h>
#include <objbase.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <audiopolicy.h>

//#define AUDCLNT_S_NO_SINGLE_PROCESS AUDCLNT_SUCCESS (0x00d)

// Remind the compiler to remind the linker to actually link the Windows libraries
#pragma comment(lib, "user32.lib")
// Needed for mecia APIs like PlaySound
#pragma comment(lib, "winmm.lib")
// Needed for system time API
#pragma comment(lib, "kernel32.lib")
// Needed for COM
#pragma comment(lib, "ole32.lib")

using namespace std;

#define LOGGING true
#define COM_AUDIO_ACTIVE true

// Declare and initialize globals
LPCTSTR sessionThreadEventName = (LPCTSTR) "AudioSessionTrackingThreadReadyEvent";
LPCTSTR processingThreadEventName = (LPCTSTR) "EventProcessingThreadReadyEvent";

//

// GetIAudioSessionManager2
// Retrieves and passes out a pointer to the IAudioSessionManager2 interface for the
// default audio endpoint device at the address pointed to by ppSessionManager..
// Returns S_OK if successful, E_POINTER if ppSessionManager is null, E_UNEXPECTED if
// the address pointed to by ppSessionManager is not empty, or the HRESULT value from
// any Windows API call which returnes a value other than S_OK (this function aborts
// at the first such occurrence).
// The caller must initialize COM in its thread before calling this function, and if
// this function returns S_OK then the caller must release the IAudioSessionManager2
// interface when no longer needed by calling its Release() method.  If this function
// does not return S_OK, then the caller does not need to clean up anything.
HRESULT GetIAudioSessionManager2(IAudioSessionManager2 ** ppSessionManager)
{
  HRESULT hr;
  IMMDevice * pDev = NULL;
  IMMDeviceEnumerator * pDevEnum = NULL;

  if(!ppSessionManager)
  {
    #if LOGGING
    printf("ERROR: In-out pointer ppSessionManager == NULL.\n");
    #endif
    return E_POINTER;
  }
  else if(*ppSessionManager)
  {
    #if LOGGING
    printf("ERROR: Passed ppSessionManager points to a non-null pointer.\n");
    #endif
    return E_UNEXPECTED;
  }

  // Initiaze the MMDevideEnumerator to get the audio endpoint device
  hr = CoCreateInstance(
      __uuidof(MMDeviceEnumerator),
      NULL,
      CLSCTX_ALL,
      __uuidof(IMMDeviceEnumerator),
      (void **) &pDevEnum
  );
  if(hr != S_OK) // If something went wrong
  {
    #if LOGGING
    printf("ERROR: CoCreateInstance failed with code %ld.\n", hr);
    #endif
    return hr;
  }
  // Get the audio endpoint device
  hr = pDevEnum -> GetDefaultAudioEndpoint(eRender, eConsole, &pDev);
  // We're done with the device enumerator, one way or the other
  pDevEnum -> Release();
  if(hr != S_OK) // If something went wrong
  {
    #if LOGGING
    printf("ERROR: GetDefaultAudioEndpoint failed with code %ld.\n", hr);
    #endif
    return hr;
  }

  hr = pDev -> Activate(
    __uuidof(IAudioSessionManager2),
    CLSCTX_ALL,
    NULL,
    (void **) ppSessionManager
  );
  pDev -> Release();
  #if LOGGING
  if(hr != S_OK)
  {
    printf("ERROR: Activate IAudioSessionManager2 failed with code %ld.", hr);
  }
  else
  {
    printf("IAudioSessionManager2 initialized.");
  }
  #endif
  return hr;
}

// Add an audio session to the programs internal tracker
// For the moment, this is a placeholder which just prints some
// information about the session
HRESULT AddAudioSession(IAudioSessionControl2 * pSession)
{
  if(!pSession)
  {
    #if LOGGING
    printf("ERROR: AddAudioSession received a null pointer.\n");
    #endif
    return E_POINTER;
  }
  HRESULT hr = S_OK;
  LPWSTR pswDisplayName = NULL;
  LPWSTR pswSessionID = NULL;
  LPWSTR pswSessionInstance = NULL;
  hr = pSession -> GetDisplayName(&pswDisplayName);
  if(hr != S_OK)
  {
    #if LOGGING
    printf("ERROR: GetDisplayName failed with error code: %ld\n", hr);
    #endif
    return hr;
  }
  hr = pSession -> GetSessionIdentifier(&pswSessionID);
  if(hr != S_OK)
  {
    #if LOGGING
    printf("ERROR: GetSessionIndentifier failed with error code: %ld\n", hr);
    #endif
    CoTaskMemFree(pswDisplayName);
    return hr;
  }
  hr = pSession -> GetSessionInstanceIdentifier(&pswSessionInstance);
  if(hr != S_OK)
  {
    #if LOGGING
    printf("ERROR: GetSessionInstanceIdentifier failed with error code: %ld\n", hr);
    #endif
    CoTaskMemFree(pswDisplayName);
    CoTaskMemFree(pswSessionID);
    return hr;
  }
  printf("Is this a double print, or is the callback firing twice?");
  printf("Audio Session found. Name: %ls, Identifier: %ls, Instance: %ls\n", pswDisplayName, pswSessionID, pswSessionInstance);
  CoTaskMemFree(pswDisplayName);
  CoTaskMemFree(pswSessionID);
  CoTaskMemFree(pswSessionInstance);
  return hr;
}

// Callback for new audio session creation
// Mostly copied from Microsoft Learn IAudioSessionNotification example
// Only the contents of OnSessionCreated have been modified
class CSessionNotifier: public IAudioSessionNotification
{
private:

    LONG             m_cRefAll;
    HWND m_hwndMain;

//    ~CSessionNotifier(){};

public:


    CSessionNotifier(HWND hWnd): 
    m_cRefAll(1),
    m_hwndMain (hWnd)

    {}

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppvInterface)  
    {    
        if (IID_IUnknown == riid)
        {
            AddRef();
            *ppvInterface = (IUnknown*)this;
        }
        else if (__uuidof(IAudioSessionNotification) == riid)
        {
            AddRef();
            *ppvInterface = (IAudioSessionNotification*)this;
        }
        else
        {
            *ppvInterface = NULL;
            return E_NOINTERFACE;
        }
        return S_OK;
    }
    
    ULONG STDMETHODCALLTYPE AddRef()
    {
        return InterlockedIncrement(&m_cRefAll);
    }
     
    ULONG STDMETHODCALLTYPE Release()
    {
        ULONG ulRef = InterlockedDecrement(&m_cRefAll);
        if (0 == ulRef)
        {
            delete this;
        }
        return ulRef;
    }

    HRESULT OnSessionCreated(IAudioSessionControl *pNewSession)
    {
        if (pNewSession)
        {
//            PostMessage(m_hwndMain, WM_SESSION_CREATED, 0, 0);
            IAudioSessionControl2 * pCtrl2 = NULL;
            HRESULT hr = S_OK;
            hr = pNewSession -> QueryInterface<IAudioSessionControl2>(&pCtrl2);
            if(hr != S_OK)
            {
              #if LOGGING
              printf("ERROR: QueryInterface for IAudioSessionControl2 failed with error code: %ld\n", hr);
              #endif
              return hr;
            }
            hr = AddAudioSession(pCtrl2);
            pCtrl2 -> Release(); // Remove this line once AddSession actually adds to a list of sessions!
            return hr;
        }
        else
        {
          return E_POINTER;
        }
    }
};

// Audio Session monitoring thread
// Populates the list of all active audio sessions and registers a callbback to add
// any new sessions created while the program is running
DWORD WINAPI AudioThreadRoutine(_In_ LPVOID pList)
{

  HANDLE hEvent = OpenEvent(SYNCHRONIZE | EVENT_MODIFY_STATE, false, sessionThreadEventName);
  if(!hEvent)
  {
    #if LOGGING
    printf("ERROR: Failed to open event handle.  Error code %ld.\n", GetLastError());
    #endif
    return (DWORD) E_FAIL;
  }

  HRESULT hr = S_OK;
  IAudioSessionManager2 * pMgr = NULL;
  IAudioSessionEnumerator * pEnum = NULL;
  CSessionNotifier sessionNotifier(NULL);
  IAudioSessionNotification * pCallback = &sessionNotifier;

  // Initialize COM for this thread
  hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
  if (hr != S_OK)
  {
    if (hr == S_FALSE) { CoUninitialize(); }
    #if LOGGING
    printf("ERROR: CoInitializeEx failed with code %ld\n", hr);
    #endif
    return 2;
  }

  // Initialize the IAudioSeesionManager2 interface
  hr = GetIAudioSessionManager2(&pMgr);
  if(hr != S_OK)
  {
    // No need for logging here - GetIAudioSessionManager2 logs its own errors.
    CoUninitialize();
    return 3;
  }

  //Register callbadk for new audio sessions
  // Do this first before going through the enumerator, in case new sessions are
  // created while processing the existing ones
  hr = pMgr -> RegisterSessionNotification(pCallback);
  if(hr != S_OK)
  {
    pMgr -> Release();
    CoUninitialize();
    return 4;
  }

  // Enumerate all of the existing sessions
  hr = pMgr -> GetSessionEnumerator(&pEnum);
  if(hr != S_OK)
  {
    #if LOGGING
    printf("ERROR: GetSessionEnumerator failed with error code %ld\n", hr);
    #endif
    pMgr -> UnregisterSessionNotification(pCallback);
    pMgr -> Release();
    CoUninitialize();
    return 5;
  }

  int numSessions = 0;
  hr = pEnum -> GetCount(&numSessions);
  if(hr != S_OK)
  {
    #if LOGGING
    printf("ERROR: Enumerator -> GetCount failed with error code: %ld\n", hr);
    #endif
    pMgr -> UnregisterSessionNotification(pCallback);
    pEnum -> Release();
    pMgr -> Release();
    CoUninitialize();
    return 6;
  }

  #if LOGGING
  printf("Preparing to review existing audio sessions. No errors yet.\n");
  #endif
  IAudioSessionControl * pCtrl = NULL;
  IAudioSessionControl2 * pCtrl2 = NULL;
  for(int i = 0; i < numSessions; i++)
  {
    hr = pEnum -> GetSession(i, &pCtrl);
    if(hr != S_OK) { break; }
    hr = pCtrl -> QueryInterface<IAudioSessionControl2>(&pCtrl2);
    pCtrl -> Release();
    if(hr != S_OK) {break;}
    
    hr = AddAudioSession(pCtrl2);

    pCtrl2 -> Release();
  }
  pEnum -> Release();

  if(hr != S_OK)
  {
    #if LOGGING
    printf("ERROR: Problem in enumeration loop, error code: %ld\n", hr);
    #endif
//    pMgr -> UnRegisterSessionNotification
    pMgr -> Release();
    CoUninitialize();
    return 7;
  }

  // Notify the main thread that we are initialized and ready to go
  SetEvent(hEvent);

  // Make sure the main thread has resumed and reset the event
  // This is ugly and theoretically unreliable, replace with something better
  Sleep(1000);

  // All done. Go to sleep and wait for end of program
  WaitForSingleObject(hEvent, INFINITE);

  pMgr -> UnregisterSessionNotification(pCallback);
  pMgr -> Release();
  CoUninitialize();
  CloseHandle(hEvent);
  return (DWORD) hr;
}



// Event procssing thread routine
// Runs in a loop and receives event reports fromthe callback in the main thread

// Callback function for the WinEvent hook
// This should be as short as possible and just gather and dispatch
// information to another thread to actually process the event, and
// try to avoid any COM calls or Windows messaging if possible to
// reduce the risk of self-interrupting reentrant execution
// The callback function should have exactly this argument list
void CALLBACK WinEventProc(
    HWINEVENTHOOK hWinEventHook,
    DWORD event,
    HWND hwnd,
    LONG idObject,
    LONG idChild,
    DWORD dwEventThread,
    DWORD dwmsEventTime
)
{
  // Check if this is a window change-of-focus event
  if (
      hwnd &&
      idObject == OBJID_WINDOW &&
      idChild == CHILDID_SELF &&
      event == EVENT_SYSTEM_FOREGROUND
  )
  {
    // Get the Process ID of the window which gained focus
    DWORD switchedProcessId;
    DWORD switchedThreadId = GetWindowThreadProcessId(hwnd, &switchedProcessId);

    // Original payload, remove when no longer needed
//    PlaySound(TEXT("C:\\Windows\\Media\\Speech Misrecognition.wav"), NULL, SND_FILENAME | SND_ASYNC);
    printf("Focus change, window of process %ld thread %ld now has focus.\n", switchedProcessId, switchedThreadId);
//    cout << "Focus change, window of prodess " << switchedProcessId << " thread " << switchedThreadId << " now has focus." << endl;
  }
}


// Main routine
// Set hook, start processor thread, run message loop, and clean up at end
// Main function name and arguments should be exactly this
int WINAPI WinMain(HINSTANCE hinst, HINSTANCE hinstPrev,
                   LPSTR lpCmdLine, int nShowCmd)
{

//  cout << "This is a test of console output." << endl;

  setvbuf(stdout, NULL, _IONBF, 0);
  HANDLE hAudioThreadEvent = CreateEvent(NULL, false, false, sessionThreadEventName);
  if(!hAudioThreadEvent)
  {
    #if LOGGING
    printf("ERROR: Creation of synchronization event failed.\n");
    #endif
    return 1;
  }
 
  // Start event processing thread here, once implemented
  
  // Start audio session tracking thread
  HANDLE hAudioThread = CreateThread(NULL, 0, AudioThreadRoutine, NULL, 0, NULL);
  if(!hAudioThread)
  {
    #if LOGGING
    printf("ERROR: Failed to start audio session tracking thread.\n");
    #endif
    return 2;
  }


  // Set the event hook for the callback function
  HWINEVENTHOOK hWinEventHook = SetWinEventHook(
     EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND,
     NULL, WinEventProc, 0, 0,
     WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);


  // Wait for other threads to finish setup. Mostly for debug purposes
  HANDLE hAudioThreadState[2] = {hAudioThreadEvent, hAudioThread};
  DWORD result = WaitForMultipleObjects(2, hAudioThreadState, false, 20000);
  if(result != WAIT_OBJECT_0)
  {
    #if LOGGING
    printf("ERROR: Failure or timeout in audio thread, return code: %ld\n", result);
    #endif
    return 3;
  }

  // Message loop, runs continuously until WM_QUIT or something goes wrong
  // GetMessage returns a "true" value normally, or a "0" on a WM_QUIT message...
  // ...or a "-1" (say WHAT, now?!) if some kind of error happens.
  // NOTE: Safer version which doesn't derail if GetMessage returns -1.
  // Doesn't attempt any special error handling, just aborts and cleans up.
  MSG msg;
  int b;
  while ((b = GetMessage(&msg, NULL, 0, 0)) && b != -1) {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }
  // Remove the event hook, if it was actually set
  SetEvent(hAudioThreadEvent);
  if (hWinEventHook) UnhookWinEvent(hWinEventHook);

  // End event procesing thread
  return 0;
}