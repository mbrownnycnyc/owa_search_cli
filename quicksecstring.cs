using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security;
using System.Runtime.InteropServices;

//this is from: http://www.codeproject.com/Tips/549109/Working-with-SecureString

namespace exch2007_owa_searcher
{

    //http://stackoverflow.com/questions/19123794/is-creating-an-explicit-or-implicit-operator-for-a-securestring-to-string-conver
    class SecureStringManager
    {

        //http://simpleprogrammer.com/2013/01/20/implicit-and-explicit-conversion-operators-in-c/
        //http://msdn.microsoft.com/en-us/library/z5z9kes2%28v=vs.71%29.aspx

        //Instantiate the secure string in your calling class as:
        // SecureString securePwd = new SecureString();
        //then set securePwd as:
        // securePwd = convertToSecureString("password");

        public SecureString convertToSecureString(string strPassword)
        {
            var secureStr = new SecureString();
            if (strPassword.Length > 0)
            {
                foreach (var c in strPassword.ToCharArray()) secureStr.AppendChar(c);
            }
            return secureStr;
        }

        public string convertToPlainTextString(SecureString secstrPassword)
        {
            IntPtr unmanagedString = IntPtr.Zero;
            try
            {
                unmanagedString = Marshal.SecureStringToGlobalAllocUnicode(secstrPassword);
                return Marshal.PtrToStringUni(unmanagedString);
            }
            finally
            {
                Marshal.ZeroFreeGlobalAllocUnicode(unmanagedString);
            }
        }

        //thanks! http://stackoverflow.com/a/3404464/843000
        public static SecureString getPasswordCLI()
        {
            SecureString pwd = new SecureString();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    Console.WriteLine("");
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length != 0)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    pwd.AppendChar(i.KeyChar);
                    Console.Write("*");
                }
            }
            return pwd;
        }

    }

}
