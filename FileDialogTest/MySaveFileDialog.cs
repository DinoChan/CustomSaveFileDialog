using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows;
using System.Windows.Interop;

namespace FileDialogTest
{
    public class MySaveFileDialog
    {
        private const int FILEBUFSIZE = 8192;

        // This is the array that stores the filename(s) the user selected in the
        // dialog box.  If Multiselect is not enabled, only the first element
        // of this array will be used.
        /// <SecurityNote>
        ///     Critical: The full file paths are critical data.
        /// </SecurityNote>
        [SecurityCritical]
        private string[] _fileNames;
        private string _initialDirectory;       // Starting directory
        private string _title;                  // Title bar of the message box
        private string _filter;                 // The file extension filters that display
                                                // in the "Files of Type" box in the dialog
        private int _filterIndex;               // The index of the currently selected
                                                // filter (a default filter index before
                                                // the dialog is called, and the filter
                                                // the user selected afterwards.)  This
                                                // index is 1-based, not 0-based.
        private string _defaultExtension;       // Extension appended first if AddExtension
                                                // is enabled
                                                //   The filter string also controls how the AddExtension feature behaves.  For
                                                //   details, see the ProcessFileNames method.
                                                /// <summary>
                                                ///       Gets or sets the current file name filter string,
                                                ///       which determines the choices that appear in the "Save as file type" or
                                                ///       "Files of type" box at the bottom of the dialog box.
                                                ///
                                                ///       This is an example filter string:
                                                ///       Filter = "Image Files(*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF|All files (*.*)|*.*"
                                                /// </summary>
                                                /// <exception cref="System.ArgumentException">
                                                ///  Thrown in the setter if the new filter string does not have an even number of tokens
                                                ///  separated by the vertical bar character '|' (that is, the new filter string is invalid.)
                                                /// </exception>
                                                /// <remarks>
                                                ///  If DereferenceLinks is true and the filter string is null, a blank
                                                ///  filter string (equivalent to "|*.*") will be automatically substituted to work
                                                ///  around the issue documented in Knowledge Base article 831559
                                                ///     Callers must have FileIOPermission(PermissionState.Unrestricted) to call this API.
                                                /// </remarks>
        public string Filter
        {
            get
            {
                // For string properties, it's important to not return null, as an empty
                // string tends to make more sense to beginning developers.
                return _filter == null ? string.Empty : _filter;
            }

            set
            {
                if (string.CompareOrdinal(value, _filter) != 0)   // different filter than what we have stored already
                {
                    string updatedFilter = value;

                    if (!string.IsNullOrEmpty(updatedFilter))
                    {
                        // Require the number of segments of the filter string to be even -
                        // in other words, there must only be matched pairs of description and
                        // file extensions.
                        //
                        // This implicitly requires there to be at least one vertical bar in
                        // the filter string - or else formats.Length will be 1, resulting in an
                        // ArgumentException.

                        string[] formats = updatedFilter.Split('|');

                        if (formats.Length % 2 != 0)
                        {
                            throw new ArgumentException("FileDialogInvalidFilter");
                        }
                    }
                    else
                    {   // catch cases like null or "" where the filter string is not invalid but
                        // also not substantive.  We set value to null so that the assignment
                        // below picks up null as the new value of _filter.
                        updatedFilter = null;
                    }

                    _filter = updatedFilter;
                }
            }
        }

        //   Using 1 as the index of the first filter entry is counterintuitive for C#/C++
        //   developers, but is a side effect of a Win32 feature that allows you to add a template
        //   filter string that is filled in when the user selects a file for future uses of the dialog.
        //   We don't support that feature, so only values >1 are valid.
        //  
        //   For details, see MSDN docs for OPENFILENAME Structure, nFilterIndex
        /// <summary>
        ///  Gets or sets the index of the filter currently selected in the file dialog box.
        ///
        ///  NOTE:  The index of the first filter entry is 1, not 0.  
        /// </summary>
        public int FilterIndex
        {
            get
            {
                return _filterIndex;
            }

            set
            {
                _filterIndex = value;
            }
        }

        /// <summary>
        ///       Gets or sets a string shown in the title bar of the file dialog.
        ///       If this property is null, a localized default from the operating
        ///       system itself will be used (typically something like "Save As" or "Open")
        /// </summary>
        /// <Remarks>
        ///     Callers must have FileIOPermission(PermissionState.Unrestricted) to call this API.
        /// </Remarks>
        /// <SecurityNote>
        ///     Critical: Do not want to allow setting the FileDialog title from a Partial Trust application.
        ///     PublicOk: Demands FileIOPermission (PermissionState.Unrestricted)
        /// </SecurityNote>
        public string Title
        {
            get
            {
                // Avoid returning a null string - return String.Empty instead.
                return _title == null ? string.Empty : _title;
            }
            [SecurityCritical]
            set
            {
                _title = value;
            }
        }

        /// <summary>
        ///  Gets or sets the initial directory displayed by the file dialog box.
        /// </summary>
        /// <Remarks>
        ///     Callers must have FileIOPermission(PermissionState.Unrestricted) to call this API.
        /// </Remarks>
        /// <SecurityNote>
        ///     Critical: Don't want to allow setting of the initial directory in Partial Trust.
        ///     PublicOk: Demands FileIOPermission (PermissionState.Unrestricted)
        /// </SecurityNote>
        public string InitialDirectory
        {
            get
            {
                // Avoid returning a null string - return String.Empty instead.
                return _initialDirectory == null ? string.Empty : _initialDirectory;
            }
            [SecurityCritical]
            set
            {
                _initialDirectory = value;
            }
        }

        /// <summary>
        ///     Gets the file names of all selected files in the dialog box.
        /// </summary>
        /// <Remarks>
        ///     Callers must have FileIOPermission(PermissionState.Unrestricted) to call this API.
        /// </Remarks>
        /// <SecurityNote> 
        ///     Critical: Do not want to allow access to raw paths to Parially Trusted Applications.
        ///     PublicOk: Demands FileIOPermission (PermissionState.Unrestricted)
        /// </SecurityNote>
        public string[] FileNames
        {
            [SecurityCritical]
            get
            {
                // FileNamesInternal is a property we use to clone
                // the string array before returning it.
                string[] files = FileNamesInternal;
                return files;
            }
        }

        //   If multiple files are selected, we only return the first filename.
        /// <summary>
        ///  Gets or sets a string containing the full path of the file selected in 
        ///  the file dialog box.
        /// </summary>
        /// <Remarks>
        ///     Callers must have FileIOPermission(PermissionState.Unrestricted) to call this API.
        /// </Remarks>
        /// <SecurityNote> 
        ///     Critical: Do not want to allow access to raw paths to Parially Trusted Applications.
        ///     PublicOk: Demands FileIOPermission (PermissionState.Unrestricted)
        /// </SecurityNote>
        public string FileName
        {
            [SecurityCritical]
            get
            {
                return CriticalFileName;
            }
            [SecurityCritical]
            set
            {

                // Allow users to set a filename to stored in _fileNames.
                // If null is passed in, we clear the entire list.
                // If we get a string, we clear the entire list and make a new one-element
                // array with the new string.
                if (value == null)
                {
                    _fileNames = null;
                }
                else
                {
                    // 

                    _fileNames = new string[] { value };
                }
            }
        }

        /// <summary>
        /// The AddExtension property attempts to determine the appropriate extension
        /// by using the selected filter.  The DefaultExt property serves as a fallback - 
        ///  if the extension cannot be determined from the filter, DefaultExt will
        /// be used instead.
        /// </summary>
        public string DefaultExt
        {
            get
            {
                // For string properties, it's important to not return null, as an empty
                // string tends to make more sense to beginning developers.
                return _defaultExtension == null ? string.Empty : _defaultExtension;
            }

            set
            {
                if (value != null)
                {
                    // Use Ordinal here as per FxCop CA1307
                    if (value.StartsWith(".", StringComparison.Ordinal)) // Allow calling code to provide 
                                                                         // extensions like ".ext" - 
                    {
                        value = value.Substring(1);    // but strip out the period to leave only "ext"
                    }
                    else if (value.Length == 0)         // Normalize empty strings to null.
                    {
                        value = null;
                    }
                }
                _defaultExtension = value;
            }
        }

        /// <summary>
        ///  Restores the current directory to its original value if the user
        ///  changed the directory while searching for files.
        ///
        ///  This property is only valid for SaveFileDialog;  it has no effect
        ///  when set on an OpenFileDialog.
        /// </summary>
        /// <Remarks>
        ///     Callers must have FileIOPermission(PermissionState.Unrestricted) to call this API.
        /// </Remarks>
        /// <SecurityNote> 
        ///     Critical: Dialog options are critical for set.
        ///     PublicOk: Demands FileIOPermission (PermissionState.Unrestricted)
        /// </SecurityNote>
        public bool RestoreDirectory
        {
            get; set;
        }

        /// <summary>
        ///  In cases where we need to return an array of strings, we return
        ///  a clone of the array.  We also need to make sure we return a 
        ///  string[0] instead of a null if we don't have any filenames.
        /// </summary>
        /// <SecurityNote>
        ///     Critical:  Accesses _fileNames, which is SecurityCritical.
        /// </SecurityNote>
        internal string[] FileNamesInternal
        {
            [SecurityCritical]
            get
            {
                if (_fileNames == null)
                {
                    return new string[0];
                }
                else
                {
                    return (string[])_fileNames.Clone();
                }
            }
        }

        //   If multiple files are selected, we only return the first filename.
        /// <summary>
        ///  Gets a string containing the full path of the file selected in 
        ///  the file dialog box.
        /// </summary>
        /// <SecurityNote> 
        ///     Critical: Do not want to allow access to raw paths to Parially Trusted Applications.
        /// </SecurityNote>
        private string CriticalFileName
        {
            [SecurityCritical]
            get
            {

                if (_fileNames == null)        // No filename stored internally...
                {
                    return string.Empty;    // So we return String.Empty
                }
                else
                {
                    // Return the first filename in the array if it is non-empty.
                    if (_fileNames[0].Length > 0)
                    {
                        return _fileNames[0];
                    }
                    else
                    {
                        return string.Empty;
                    }
                }
            }
        }

        public bool? ShowDialog()
        {
            var openFileName = new OpenFileName();
            Window window = Application.Current.Windows.OfType<Window>().Where(w => w.IsActive).FirstOrDefault();
            if (window != null)
            {
                var wih = new WindowInteropHelper(window);
                IntPtr hWnd = wih.Handle;
                openFileName.hwndOwner = hWnd;
            }

            openFileName.structSize = Marshal.SizeOf(openFileName);
            openFileName.filter = MakeFilterString(Filter);
            openFileName.filterIndex = FilterIndex;
            openFileName.fileTitle = new string(new char[64]);
            openFileName.maxFileTitle = openFileName.fileTitle.Length;
            openFileName.initialDir = InitialDirectory;
            openFileName.title = Title;
            openFileName.defExt = DefaultExt;
            openFileName.structSize = Marshal.SizeOf(openFileName);
            openFileName.flags |= FOS.NOTESTFILECREATE | FOS.OVERWRITEPROMPT;
            if (RestoreDirectory)
                openFileName.flags |= FOS.NOCHANGEDIR;


            // lpstrFile
            // Pointer to a buffer used to store filenames.  When initializing the
            // dialog, this name is used as an initial value in the File Name edit
            // control.  When files are selected and the function returns, the buffer
            // contains the full path to every file selected.
            char[] chars = new char[FILEBUFSIZE];

            for (int i = 0; i < FileName.Length; i++)
            {
                chars[i] = FileName[i];
            }
            openFileName.file = new string(chars);
            // nMaxFile
            // Size of the lpstrFile buffer in number of Unicode characters.
            openFileName.maxFile = FILEBUFSIZE;

            if (LibWrap.GetSaveFileName(openFileName))
            {
                FileName = openFileName.file;
                return true;
            }
            return false;
        }



        /// <summary>
        ///     Converts the given filter string to the format required in an OPENFILENAME_I
        ///     structure.
        /// </summary>
        private static string MakeFilterString(string s, bool dereferenceLinks = true)
        {
            if (string.IsNullOrEmpty(s))
            {
                // Workaround for VSWhidbey bug #95338 (carried over from Microsoft implementation)
                // Apparently, when filter is null, the common dialogs in Windows XP will not dereference
                // links properly.  The work around is to provide a default filter;  " |*.*" is used to 
                // avoid localization issues from description text.
                //
                // This behavior is now documented in MSDN on the OPENFILENAME structure, so I don't
                // expect it to change anytime soon.
                if (dereferenceLinks && System.Environment.OSVersion.Version.Major >= 5)
                {
                    s = " |*.*";
                }
                else
                {
                    // Even if we don't need the bug workaround, change empty
                    // strings into null strings.
                    return null;
                }
            }

            StringBuilder nullSeparatedFilter = new StringBuilder(s);

            // Replace the vertical bar with a null to conform to the Windows
            // filter string format requirements
            nullSeparatedFilter.Replace('|', '\0');

            // Append two nulls at the end
            nullSeparatedFilter.Append('\0');
            nullSeparatedFilter.Append('\0');

            // Return the results as a string.
            return nullSeparatedFilter.ToString();
        }

    }

    internal class FOS
    {
        public const int OVERWRITEPROMPT = 0x00000002;
        public const int STRICTFILETYPES = 0x00000004;
        public const int NOCHANGEDIR = 0x00000008;
        public const int PICKFOLDERS = 0x00000020;
        public const int FORCEFILESYSTEM = 0x00000040;
        public const int ALLNONSTORAGEITEMS = 0x00000080;
        public const int NOVALIDATE = 0x00000100;
        public const int ALLOWMULTISELECT = 0x00000200;
        public const int PATHMUSTEXIST = 0x00000800;
        public const int FILEMUSTEXIST = 0x00001000;
        public const int CREATEPROMPT = 0x00002000;
        public const int SHAREAWARE = 0x00004000;
        public const int NOREADONLYRETURN = 0x00008000;
        public const int NOTESTFILECREATE = 0x00010000;
        public const int HIDEMRUPLACES = 0x00020000;
        public const int HIDEPINNEDPLACES = 0x00040000;
        public const int NODEREFERENCELINKS = 0x00100000;
        public const int DONTADDTORECENT = 0x02000000;
        public const int FORCESHOWHIDDEN = 0x10000000;
        public const int DEFAULTNOMINIMODE = 0x20000000;
        public const int FORCEPREVIEWPANEON = 0x40000000;
    }


    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public class OpenFileName
    {
        internal int structSize = 0;
        internal IntPtr hwndOwner = IntPtr.Zero;
        internal IntPtr hInstance = IntPtr.Zero;
        internal string filter = null;
        internal string custFilter = null;
        internal int custFilterMax = 0;
        internal int filterIndex = 0;
        internal string file = null;
        internal int maxFile = 0;
        internal string fileTitle = null;
        internal int maxFileTitle = 0;
        internal string initialDir = null;
        internal string title = null;
        internal int flags = 0;
        internal short fileOffset = 0;
        internal short fileExtMax = 0;
        internal string defExt = null;
        internal int custData = 0;
        internal IntPtr pHook = IntPtr.Zero;
        internal string template = null;
    }

    public class LibWrap
    {
        // Declare a managed prototype for the unmanaged function. 
        [DllImport("Comdlg32.dll", SetLastError = true, ThrowOnUnmappableChar = true, CharSet = CharSet.Auto)]
        public static extern bool GetSaveFileName([In, Out] OpenFileName ofn);
    }

}
