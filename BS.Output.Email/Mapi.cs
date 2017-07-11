using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace BS.Output.Email
{

  internal class Mapi
  {

    [DllImport("MAPI32.DLL")]
    internal static extern int MAPISendMail(IntPtr session, IntPtr hwnd, MapiMessage message, int flg, int rsv);

    public void SendMail(string filePath)
    {
      
      // Alloc file
      IntPtr objIntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(MapiFileDescriptor)));
      MapiFileDescriptor objMapiFileDescriptor = new MapiFileDescriptor();
      objMapiFileDescriptor.position = -1;
      objMapiFileDescriptor.name = System.IO.Path.GetFileName(filePath);
      objMapiFileDescriptor.path = filePath;
      Marshal.StructureToPtr(objMapiFileDescriptor, objIntPtr, false);

      // Send Mail    
      MapiMessage objMessage = new MapiMessage();
      objMessage.FileCount = 1;
      objMessage.Files = objIntPtr;
      MAPISendMail(IntPtr.Zero, IntPtr.Zero, objMessage, 8, 0);

      // Dealoc file
      Marshal.DestroyStructure(objMessage.Files, typeof(MapiFileDescriptor));
      Marshal.FreeHGlobal(objMessage.Files);
        
    }
    
  }

  [StructLayout(LayoutKind.Sequential)]
  internal class MapiFileDescriptor
  {
    public int reserved;
    public int flags;
    public int position;
    public string path;
    public string name;
    public IntPtr type = IntPtr.Zero;
  }

  [StructLayout(LayoutKind.Sequential)]
  internal class MapiMessage
  {
    public int Reserved;
    public string Subject;
    public string NoteText;
    public string MessageType;
    public string DateReceived;
    public string ConversationID;
    public int Flags;
    public IntPtr Originator = IntPtr.Zero;
    public int RecipientCount;
    public IntPtr Recipients = IntPtr.Zero;
    public int FileCount;
    public IntPtr Files = IntPtr.Zero;
  }

}
