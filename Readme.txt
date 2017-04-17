*** Interprocess Communication via Semaphores
*** Originally released December 2010 by Joe Jordan, Ignite Software Inc.

Ah, semaphores, the staple of any good operating system. I searched for a VB6 example implementing the semaphore functions and came up mostly empty-handed. This class attempts to fill that missing gap in the world of VB6 examples.

In developing this class, I initially thought it would be simple to create a global semaphore that all users would have access to. After all, the documentation states that: "The semaphore name can have a "Global\" or "Local\" prefix to explicitly create the object in the global or session name space." Little did I know that I would have to delve into the depths of Windows security functions in order to actually provide *true* Global semaphore functionality. After many failed attempts, approaches and cryptic error messages (The revision level is unknown. ***?), I believe that the infamous ACL dragon has, for our intentions at least, been slayed.

I took the advice of one Anne Gunn and implemented some additional security for our global semaphore, so rogue applications can't steal our lunch money completely.

The majority of the trial and error took place in finding the proper way to call and declare the security APIs. I thought I was 98% done, so I tested on XP to see if it worked there, as I figured if I could get it to work on Windows 7-64 bit, surely it would work in the UAC-less environment of XP. Well, it worked fine in the IDE, but spit out an invalid memory access error when compiled. I had gotten a similar error while testing in Windows 7 and tracked it down to using the actual struct when calling CreateSemaphore rather than the pointer. So I had to go back through each call and test to see which one needed the actual struct instead of the pointer. Turns out it was SetSecurityDescriptorDacl that needed to accept an actual SECURITY_DESCRIPTOR rather than a pointer to one. After the 2nd such discovery, I went back and used the actual structs whenever possible as a precaution.

The class was lightly tested on Windows 2000, XP, Vista and 7. If you come across any issues or have any improvements or suggestions please let me know.
---
FYI: Viewing your Semaphore

I know that for many, semaphores are an unfamiliar topic, but they're really just simple communication devices. The sample application gives a good rundown of the features available, but to actually see what's going on, you'll need special software, specifically Process Explorer. A semaphore is one of many types of handles available in Windows programs, so we'll have to enable handle viewing in Process Explorer.

1. Download and run Process Explorer
2. Click on View -> Show Lower Pane
3. Click on View -> Lower Pane View -> Handles
4. Select VB6 or your compiled SemaphoreTest app; all handles associated with the EXE will be listed. I like to sort the Type column so Semaphores are near the top.
5. Select TestSemaphore1 and right-click -> Properties. ProcExp will give you all the info you'd ever want to know about your semaphore.

http://i.imgur.com/hKVaD.png

---
CHANGELOG:

v1.2.1
- Added optional DecrementSize to the Decrement function to reduce the semaphore value a custom amount
- Cleaned up a few functions, removed some unnecessary variables and declares

v1.2
- Fixed and Finished QueryHandleCount function
- Added (currently unused) code to convert LARGE_INTEGER to VB Date, in case one day Windows decides to start returning CreationTime for semaphores
- Removed VirtualAlloc declares and constants; using all Heap functions now

v1.1
+ Added IsSemaphore function to quickly test for the existence of a semaphore
+ Added HandleSecurity feature which prevents the closing of the semaphore handle via CloseHandle
+ Added caching of ValidateDLL result to improve speed
+ Started Adding QueryHandleCount function; if anyone can fix it please let me know
- Fixed logic error in getting the semaphore global state if we opened an existing semaphore on initialize call

v1.0
- Initial Release

----
Credits:
http://undocumented.ntinternals.net/ for information on the undocumented NtQuerySemaphore API function and the (also undocumented) SEMAPHORE_QUERY_STATE permission constant.

Anne Gunn for her excellent, thorough and well written article and accompanying code on creating a not-quite-null dacl, and explaining the benefits of doing so.
http://www.codeguru.com/cpp/w-p/win32/tutorials/article.php/c4545

Matts_User_Name of the SysInternals forums for the QueryName function and for helping debug QueryHandleCount.
http://forum.sysinternals.com/handle-name-help-ntqueryObject_topic14435_page2.html

IrfanAhmad on the MSDN forums for his thread on how to share a semaphore:
http://social.msdn.microsoft.com/Forums/en/windowssdk/thread/335db156-b1f7-45e2-b3d1-f0e79e386744