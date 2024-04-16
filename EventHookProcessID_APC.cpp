// Modified window event hook program, based on example from R. Chen
// Makes a text log entry whenever a task switch occurs, noting the new process of
// the window which gained focus, the presumed process of the presumed previous
// window which lost focus, and the time of the event.


/*
 * Basic principle of the program involves three parts
 * - Event hook which listens for window-change events, checks that the event is valid, identifies
 *     the process which gained focus, and either processes or (more likely) dispatches the event
 * - Main method which initializes any other threads or resources, sets the event hook, loops and
 *     pumps the message queue for the event hook, and cleans up everything at the end
 * - Processing routine which performs the auto mute and unmute of the switched program
 *   - Check that the new process id is valid and should not be ignored
 *   - Get pointers to the audio session control for the old and new process
 *   - Use the pointers to mute the old process and unmute the new one
 *   - This requires using COM calls which could indirectly trigger GetMessage in the calling thread, so
 *      it is safer to do this in a different thread than the event hook to avoid possible re-entrancy
 * 
 * The process losing focus cannot easily be directly identified, so this program relies on remembering
 * the previous process to gain focus and assuming continuity (i.e. that no undetected focus changes
 * occur in between the detected ones) to identify the process which should be muted.  This requires
 * access to state data in the mute/unmute routine, and if using a separate thread for COM calls to
 * avoid re-entrancy in the event hook then a way to pass asynchronous FIFO messages to the processing
 * thread is needed, either of which is messy since neither the event hook callback nor the Windows APC
 * routine callback allows a USERDATA pointer or similar as an argument.  This leave several possible
 * approaches:
 * - Single thread, everything in the event hook (Re-entrancy hazard)
 * - Use globals + an APC routine in a second thread
 * - Use globals + queue + waitable synchronization Events (essentially APC by hand)
 * - Use globals + custom events in the Windows message queue for the second thread
 * - Use a mailslot or a named pipe to pass message to the second thread (low performance?)
 * - Use one global (for memory) + thread pool routines (which allow USERDATA in the callback)
 * equivalent to "").
 * 
 * 
 */
// Standard C++ header files
#include <iostream>
#include <thread>
#include <queue>

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
// #include <mmsystem.h>
//#include <combaseapi.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <audiopolicy.h>

#define AUDCLNT_S_NO_SINGLE_PROCESS AUDCLNT_SUCCESS (0x00d)

// Remind the compiler to remind the linker to actually link the Windows libraries
#pragma comment(lib, "user32.lib")
// Needed for media APIs like PlaysSound
#pragma comment(lib, "winmm.lib")
// Needed for system time API
#pragma comment(lib, "kernel32.lib")
// Needed for COM
#pragma comment(lib, "ole32.lib")

#define LOGGING false
#define USING_SINGLE_THREAD_AND_GLOBAL false
#define USING_APC_THREAD_AND_GLOBAL true
#define COM_AUDIO_ACTIVE true

// Declare and initialize a thread-safe FIFO structure here
//

// GetIAudioSessionManager2
// Retrieves and passes out a pointer to the IAudioSessionManager2 interface for the
// default audio endpoint device at the address pointed to by ppSessionManager,
// overwriting any existint pointer at that address (may cause memory leaks or
// undefined behavior if that address contains a non-null pointer).
// Returns S_OK if successful, E_POINTER if ppSessionManager is null, or the HRESULT
// value from any Windows API call which returned a value other than S_OK (this
// function aborts at the first such occurrence).
// The caller must initialize COM in its thread before calling this function, and if
// this function returns S_OK then the caller must release the IAudioSessionManager2
// interface when no longer needed by calling its Release() method.  If this function
// returns any value other than S_OK, then the dereferenced address *ppSessionManager
// will contain a null pointer (unless ppSessionManager itself is null), and the
// caller does not need to clean up anything.
HRESULT GetIAudioSessionManager2(IAudioSessionManager2 ** ppSessionManager)
{
  HRESULT hr;
  IMMDevice * pDev = NULL;
  IMMDeviceEnumerator * pDevEnum = NULL;

  // Make sure the passed in-out ptr for the session manager is not NULL, 
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
    printf("WARNING: Passed ppSessionManager points to a non-null pointer.\n");
    printf("Attempting to release the old interface before overwriting.\n");
    #endif
    (*ppSessionManager) -> Release();
    *ppSessionManager = NULL;
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
    printf("ERROR: CoCreateInstance failed with code %d.\n", hr);
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
    printf("ERROR: GetDefaultAudioEndpoint failed with code %d.\n", hr);
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
    printf("ERROR: Activate IAudioSessionManager2 failed with code %d.", hr);
  }
  else
  {
    printf("IAudioSessionManager2 initialized.");
  }
  #endif
  return hr;
}

// Auto Mute handler method
// Enumerates all sound sessions on the system, and mutes any single-process session
// associated with prevProcId and un-mutes any single-process session associated with
// newProcId.  Does not alter the state of any cross-process sound session or system
// events sound session.  Returns S_OK if successful, or aborts immediately if any
// Windows API call fails and returns the HRESULT from that failure.  Does nothing and
// immediately returns E_POINTER if pSessionMgr is a null pointer, E_INVALIDARG if
// either newProcId or prevProcId is zero, or S_FALSE if newProcId and prevProcId are
// the same (and not zero), in that order; note that the former cases indicate failure
// due to bad input, while the last represents a special case of valid input.
HRESULT AutoMuteRoutine(
    IAudioSessionManager2 * pSessionMgr,
    DWORD newProcId,
    DWORD prevProcId
)
{
  #if LOGGING
  printf("Switch event received.  Process %d to %d.\n", prevProcId, newProcId);
  #endif

  // If switching windows from the same process, do nothing
  if(!newProcId || !prevProcId)
  {
    #if LOGGING
    printf("ERROR: Invalid process id.  This should never be zero.\n");
    #endif
    return E_INVALIDARG;
  }
  if(newProcId == prevProcId) {
    #ifdef LOGGING
    printf("INFO: Both windows belong to same process.  No action.");
    #endif
    return S_FALSE;
  }

  HRESULT hr = S_OK;
  IAudioSessionEnumerator * pEnum = NULL;
  int numSessions = 0;

  #if COM_AUDIO_ACTIVE
  if(!pSessionMgr)
  {
    #if LOGGING
    printf("ERROR: Received null pointer for IAudioSessionManager2.\n");
    #endif
    return E_POINTER;
  }
  hr = pSessionMgr -> GetSessionEnumerator(&pEnum);
  if(hr != S_OK)
  {
    #if LOGGING
    printf("ERROR: GetSessionEnumerator failed with code %d\n", hr);
    #endif
    return hr;
  }

  hr = pEnum -> GetCount(&numSessions);
  if(hr != S_OK)
  {
    #if LOGGING
    printf("ERROR: SessionEnumerator GetCount failed with code %d\n", hr);
    #endif
    pEnum -> Release();
    return hr;
  }

  #if LOGGING
  printf("Retrieved Session Enumerator.  %d audio sessions.\n", numSessions);
  #endif

  for(int i = 0; i < numSessions; i++)
  {
    IAudioSessionControl * pCtrl = NULL;
    IAudioSessionControl2 * pCtrl2 = NULL;
    ISimpleAudioVolume * pVol = NULL;
    DWORD procId = NULL;

    #if LOGGING && 0x02
    printf("Session number %d out of %d.\n", i, numSessions);
    #endif

    hr = pEnum -> GetSession(i, &pCtrl);
    if(hr != S_OK)
    {
      #if LOGGING
      printf("ERROR: Session %d - GetSession failed for with code %d\n", i, hr);
      #endif
      break;
    }
    
    hr = pCtrl -> QueryInterface<IAudioSessionControl2>(&pCtrl2);
    pCtrl -> Release(); // Don't need this any more, one way or the other
    if(hr != S_OK)
    {
      #if LOGGING && 0x02
      printf("ERROR: Session %d - QueryInterface for IAudioSessionControl2 failed with code %d\n", i, hr);
      #endif
      break;
    }

    // Do not auto-mute system notification session
    // This is a normal occurrence.  Just skip this session.
    hr = pCtrl2 -> IsSystemSoundsSession();
    if(hr == S_OK)
    {
      pCtrl2 -> Release();
      #if LOGGING
      printf("INFO: Session %d - This is a system sounds session.  Skip to next.\n", i);
      #endif
      continue;
    }
    else if(hr != S_FALSE) // This, on the other hand, is NOT normal
    {
      pCtrl2 -> Release();
      #if LOGGING
      printf("ERROR: Session %d - IsSystemSoundsSession failed with code %d\n", i, hr);
      #endif
      break;
    }

    hr = pCtrl2 -> GetProcessId(&procId);
      
    // If this is a cross-process session, just skip it
    // This is not a failure, so set hr = S_OK so that the
    // outer loop doesn't mistake this for an abort if there
    // are no more sessions in the list after this
    // This may change in later versions
    if(hr == AUDCLNT_S_NO_SINGLE_PROCESS)
    {
      pCtrl2 -> Release();
      hr = S_OK; // So the main loop doesn't panic, if this was the last session
      #if LOGGING
      printf("INFO: Session %d - This is a cross-process session.  Skip to next.", i);
      #endif
      continue;
    }
    else if(hr != S_OK)
    {
      pCtrl2 -> Release();
      #if LOGGING
      printf("ERROR: Session %d - GetProcessId failed with code %d\n", i, hr);
      #endif
      break;
    }
    
    if(procId == newProcId)
    {
      #if LOGGING
      printf("New process session found.  Sesiion %d, process id %d\n", i, procId);
      #endif

      hr = pCtrl2 -> QueryInterface<ISimpleAudioVolume>(&pVol);
      pCtrl2 -> Release(); // Release the moment it isn't needed
      if(hr != S_OK)
      {
        #if LOGGING
        printf("ERROR: Session %d - QueryInterface for ISimpleAudioVolume failed with code %d\n", i, hr);
        #endif
        break;
      }

      hr = pVol -> SetMute(false, NULL);
      pVol -> Release();
      if(hr != S_OK)
      {
        #if LOGGING
        printf("ERROR: Session %d - SetMute failed with code %d\n", i, hr);
        #endif
        break;
      }
    }
    else if(procId == prevProcId)
    {
      #if LOGGING
      printf("Old process session found.  Session %d, process id %d.\n", i, procId);
      #endif

      hr = pCtrl2 -> QueryInterface<ISimpleAudioVolume>(&pVol);
      pCtrl2 -> Release(); // Release the second it isn't needed
      if(hr != S_OK)
      {
        #if LOGGING
        printf("ERROR: Session %d - QueryInterface for ISimpleAudioVolume failed with code %d\n", i, hr);
        #endif
        break;
      }

      hr = pVol -> SetMute(true, NULL);
      pVol -> Release();
      if(hr != S_OK)
      {
        #if LOGGING
        printf("ERROR: Session %d - SetMute failed with codde %d\n", i, hr);
        #endif
        break;
      }
    }
    else
    {
      pCtrl2 -> Release();
    }
  }

  // Release the IAudioSessionEnumerator interface
  pEnum -> Release();

  // hr may be AUDCLNT_S_NO_SINGLE_PROCESS rather than S_OK if
  // the the last session was a cross-process session, which
  // should be be returned as S_OK since this is not a failure.
  if(hr == AUDCLNT_S_NO_SINGLE_PROCESS) {hr = S_OK;}
  #endif

  return hr;
}



// Event procssing thread routine
// Runs in a loop and receives event reports fromthe callback in the main thread
DWORD WINAPI EventThreadRoutine(_In_ LPVOID pQueue)
{
  // // Do initial setup and variable declarations before the loop
  DWORD newProcId = NULL;
  DWORD prevProcId = GetCurrentProcessId();

  #if LOGGING
  // // Open log file, if used
  //
  #endif

  if(!pQueue)
  {
    #if LOGGING
    printf("ERROR: Received null pointer for event queue.\n");
    #endif
    return (DWORD) E_POINTER;
  }

  LPCTSTR eventName = (LPCTSTR) "processingThreadEvent";

  HANDLE hEvent = OpenEvent(SYNCHRONIZE | EVENT_MODIFY_STATE, false, eventName);
  if(!hEvent)
  {
    #if LOGGING
    printf("ERROR: Failed to open event handle.  Error coed %d.\n", GetLastError());
    #endif
    return (DWORD) E_FAIL;
  }

  HRESULT hr = NULL;
  IAudioSessionManager2 * pMgr = NULL;

  #if COM_AUDIO_ACTIVE
  // Initialize COM for this thread
  hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
  if (hr != S_OK)
  {
    if (hr == S_FALSE) { CoUninitialize(); }
    #if LOGGING
    printf("ERRoR: CoInitializeEx failed with code %d\n", hr);
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
  #endif

  // Notify the main thread that we are initialized and ready to go
  SetEvent(hEvent);

  // Main event processing loop
  // Need to replace this with an exit condition
  while(1)
  {
    // Wait for an event from the queue
    WaitForSingleObject(hEvent, INFINITE);

    // Retrieve the next event from the queue
    //

    // At this point, newProcId has the id of the new process

    // Normal exit condition.  Main thread sends this to quit.
    if(newProcId == -1){ hr = S_OK; break; }

    hr = AutoMuteRoutine(pMgr, newProcId, prevProcId);

    if((hr != S_OK) && (hr != S_FALSE))
    {
      #if LOGGING
      printf("ERROR: Auto Mute Routine failed.  Exiting program.\n");
      #endif
      break;
    }

    // Set the new process id as the previous process id for the next event
    prevProcId = newProcId;
  }

  #if COM_AUDIO_ACTIVE
  pMgr -> Release();
  CoUninitialize();
  CloseHandle(hEvent);
  #endif

  // Close log file here, if used
  //

  return (DWORD) hr;
}

#if USING_APC_THREAD_AND_GLOBAL
HANDLE ghAPCThread = NULL;
DWORD gPrevProcessId = NULL;
IAudioSessionManager2 * gpSessionMgr = NULL;
boolean gbEndProgram = false;
//PAPCFUNC APCAutoMuteRoutine;
//PAPCFUNC APCThreadExit;

void APCAutoMuteRoutine(ULONG_PTR pData)
{
  DWORD newProcId = *((DWORD *) pData);
  if(newProcId == gPrevProcessId)
  {
    #if LOGGING
    printf("New and old window belong to same process.  Do nothing.\n");
    #endif
    return;
  }

  DWORD prevProcId = gPrevProcessId;
  gPrevProcessId = newProcId;

  HRESULT hr = AutoMuteRoutine(gpSessionMgr, newProcId, prevProcId);
  if(hr != S_OK)
  {
    gbEndProgram = true;
    #if LOGGING
    printf("Auto mute routine failed.  Aborting program.\n");
    #endif
    gpSessionMgr -> Release();
    CoUninitialize();
    ExitThread((DWORD) hr);
  }
  return;
}

// APC routine which cleans up and terminates the APC thread.
// The argument pExitCode can be NULL for a normal exit (with
// exit code 0), or a pointer to a DWORD thread exit value.
void APCThreadExit(LPVOID pExitCode)
{
  DWORD exitCode = 0;
  if(pExitCode) { exitCode = *((DWORD *) pExitCode); }
  #if LOGGING
  if(exitCode)
  {
    printf("APC thread ordered to terminate with exit code %d.\n", exitCode);
  else
  {
    printf("APC thread terminated normally.  (Exit code %d)\n", exitCode);
  }
  #endif
  gpSessionMgr -> Release();
  gpSessionMgr = NULL;
  CoUninitialize();
  ExitThread(exitCode);
}

// Thread routine for processing APC events.
// Initializes COM and globals and then loops continuously
// in an alertable wait state to receive APC requests.
DWORD WINAPI APCThreadRoutine(LPVOID pData)
{
  #if COM_AUDIO_ACTIVE
  HRESULT hr = S_OK;
  hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
  if(hr != S_OK)
  {
    #if LOGGING
    printf("CoInitializeEx failed with code %d.\n", hr);
    #endif
    if(hr = S_FALSE) { CoUninitialize(); hr = E_FAIL; }
    return (DWORD) hr;
  }
  hr = GetIAudioSessionManager2(&gpSessionMgr);
  if(hr != S_OK) { return hr; }
  #endif

  #if LOGGING
  printf("APC thread ready.");
  #endif

  while(!gbEndProgram) { SleepEx(10000, true); }

  #if COM_AUDIO_ACTIVE
  gpSessionMgr -> Release();
  gpSessionMgr = NULL;
  CoUninitialize();
  #endif

  return (DWORD) gbEndProgram;
}

#endif

#if USING_SINGLE_THREAD_AND_GLOBAL
DWORD gPrevProcessId = NULL;
IAudioSessionManager2 * gpSessionMgr = NULL;
boolean gbEndProgram = false;
int reentrancyCount = 0;
#endif

// Callback function for the WinEvent hook
// This should be as short as possible and just gather and dispatch
// data to another thread to be processed, and try to avoid any COM
// calls or Windows messaging if possible to avoid risk of re-entrance
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
  #if USING_APC_THREAD_AND_GLOBAL || USING_SINGLE_THREAD_AND_GLOBAL
  if(gbEndProgram) { PostQuitMessage((DWORD) E_ABORT); return; }
  #endif
  // Check if this is a window change-of-focus event
  // the "if(hwnd && ...)" is effectively "if(hwnd != NULL && ...)"
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
    PlaySound(TEXT("C:\\Windows\\Media\\Speech Misrecognition.wav"),
             NULL, SND_FILENAME | SND_ASYNC);
    // Dispatch the process ID to a thread-safe FIFO data structure
    //
    // // DispatchToQueue(switchedProcessID, switchedThreadID)

    #if USING_SINGLE_THREAD_AND_GLOBAL
    if(switchedProcessId == gPrevProcessId)
    {
      #if LOGGING
      printf("New and old window belong to same process. Do nothing.\n");
      #endif
      return;
    }
    // Retrieve and update previous process id before any possible re-entrancy
    DWORD prevProcId = gPrevProcessId;
    gPrevProcessId = switchedProcessId;
    // Check for potential re-entrancy
    if(reentrancyCount > 0)
    {
      #if LOGGING
      printf("WARNING: Re-entrant callback.  Re-entered %d time(s).\n", reentrancyCount);
      #endif
      #if (ALLOW_CALLBACK_REENTRANCE == -1)
      // Do something ?
      #else
      if(reentrancyCount > ALLOW_CALLBACK_REENTRANCE)
      {
        gbEndProgram = true;
        #if LOGGING
        printf("ERROR: Re-entrance test limit exceeded.  Aborting program.\n");
        #endif
        PostQuitMessage((int) E_ABORT);
        return;
      }
      #endif
    }
    reentrancyCount ++;
    // Operations after this may risk re-entrancy
    HRESULT hr = AutoMuteRoutine(gpSessionMgr, switchedProcessId, prevProcId);
    // End of operations which may risk re-entrancy
    reentrancyCount --;
    if((hr != S_OK) && (hr != S_FALSE))
    {
      gbEndProgram = true;
      PostQuitMessage((int) hr);
      return;
    }
    #endif

    #if USING_APC_THREAD_AND_GLOBAL
      if(!QueueUserAPC(APCAutoMuteRoutine, hAPCThread, (ULONG_PTR) &switchedProcessId))
      {
        gbEndProgram = true;
        PostQuitMessage((DWORD) E_FAIL);
        #if LOGGING
        printf("QueueUserAPC failed.  Aborting program.\n");
        #endif
        return;
      }
    #endif
  }  
}

// Main routine
// Set hook, start processor thread, run message loop, and clean up at end
// Main function name and arguments should be exactly this
int WINAPI WinMain(HINSTANCE hinst, HINSTANCE hinstPrev,
                   LPSTR lpCmdLine, int nShowCmd)
{
  // // Create / open the log file for the callback method here?
  //

  // // Create the event queue object here?
  //
  DWORD exitCode = 0;

  #if USING_SINGLE_THREAD_AND_GLOBAL
  HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
  if(hr != S_OK)
  {
    if(hr == S_FALSE) { CoUninitialize(); }
    #if LOGGING
    printf("COM Initialization falied with code %d.\n", hr);
    #endif
  }
  else
  {
    hr = GetIAudioSessionManager2(&gpSessionMgr);
    if(hr != S_OK)
    {
      #if LOGGING
      printf("GetIAudioSessionManager2 failed with code %d.\n", hr);
      #endif
      CoUninitialize();
    }
  }
  if(hr != S_OK)
  {
    // close log file, if used
    //
    return (int) hr;
  }
  gPrevProcesId = GetCurrentProcessId();
  #endif

  #if USING_APC_THREAD_AND_GLOBAL
  ghAPCThread = CreateThread(NULL, 0, APCProcEvents, NULL, 0, NULL);
  if(!ghAPCThread)
  {
    #if LOGGING
    printf("APC Thread failed to start.  Aborting program.\n");
    #endif
    return (int) E_ABORT;
  }
  // wait for thread to initialize
  //
  gPrevProcessId = GetCurrentProcessId();
  #endif

  // // Start thread for routine to actually process the change-of-focus events
  LPCTSTR eventName = (LPCTSTR) "threadEvent";
  HANDLE hEvent = CreateEvent(NULL, false, false, eventName);
  if(!hEvent)
  {
    #if LOGGING
    printf("Failed to create event object.  Error code %d.\n", GetLastError());
    #endif
    return (DWORD) E_FAIL;
  }

  // Third argument is the thread method, fourth is the userData pointer
  HANDLE hEventThread = CreateThread(NULL, 0, EventThreadRoutine, NULL, 0, NULL);
  if(!hEventThread)
  {
    #if LOGGING
    printf("Failed to start event handling thread.  Error code %d.\n", GetLastError());
    #endif
    CloseHandle(hEvent);
    return (DWORD) E_FAIL;
  }

  HANDLE waitArray[] = {hEvent, hEventThread};
  exitCode = WaitForMultipleObjects(2, waitArray, false, INFINITE);
  if(exitCode != WAIT_OBJECT_0)
  {
    #if LOGGING
    printf("Processing thread initialization failed.\n");
    #endif
    CloseHandle(hEventThread);
    CloseHandle(hEvent);
    ExitProcess(1);
  }

  // Set the event hook for the callback function
  HWINEVENTHOOK hWinEventHook = SetWinEventHook(
     EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND,
     NULL, WinEventProc, 0, 0,
     WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);

  // Message loop, runs continuously until WM_QUIT or something goes wrong
  // GetMessage returns a "true" value normally, or a "0" on a WM_QUIT message...
  // ...or a "-1" (say WHAT now?!) if some kind of error happens.
  // Doesn't attempt any special error handling, just aborts and cleans up.
  MSG msg;
  boolean b;
  while ((b = GetMessage(&msg, NULL, 0, 0)) && b != -1)
  {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }

  // Remove the event hook, if it was actually set
  if (hWinEventHook) UnhookWinEvent(hWinEventHook);
  
  #if USING_SINGLE_THREAD_AND_GLOBAL
  gpSessionMgr -> Release();
  CoUninitialize();
  if((b == 0) && !gEndProgram)
  {
    exitCode = (int) S_OK;
  }
  else
  {
    exitCode = (int) msg -> lparam;
  }
  #endif

  #if USING_APC_THREAD_AND_GLOBAL
  // Send quit message to APC thread, just to be safe
  //
  b = QueueUserAPC(APCThreadExit, ghAPCThread, NULL);
  exitCode = WaitForSingleObject(ghAPCThread, INFINITE);
  CloseHandle(ghAPCThread);
  #endif

  // Send shutdown message to processing thread, and wait for it to end
  //
  // send shutdown message
  //

  exitCode = WaitForSingleObject(hEventThread, INFINITE);
  b = CloseHandle(hEventThread);
  b = CloseHandle(hEvent);

  // Close the log file, if used
  //

  return (int) exitCode;
}