// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Collections.Generic;
using System.Globalization;
using System.Text;

using System.Management.Automation.Internal;

namespace System.Management.Automation
{
    internal static class EncodingConversion
    {
        internal const string ANSI = "ansi";
        internal const string Ascii = "ascii";
        internal const string BigEndianUnicode = "bigendianunicode";
        internal const string BigEndianUtf32 = "bigendianutf32";
        internal const string Default = "default";
        internal const string OEM = "oem";
        internal const string String = "string";
        internal const string Unicode = "unicode";
        internal const string Unknown = "unknown";
        internal const string Utf7 = "utf7";
        internal const string Utf8 = "utf8";
        internal const string Utf8Bom = "utf8BOM";
        internal const string Utf8NoBom = "utf8NoBOM";
        internal const string Utf32 = "utf32";

        internal static readonly string[] TabCompletionResults = {
                ANSI, Ascii, BigEndianUnicode, BigEndianUtf32, OEM, Unicode, Utf7, Utf8, Utf8Bom, Utf8NoBom, Utf32
            };

        internal static readonly Dictionary<string, Encoding> encodingMap = new(StringComparer.OrdinalIgnoreCase)
        {
            { ANSI, Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.ANSICodePage) },
            { Ascii, Encoding.ASCII },
            { BigEndianUnicode, Encoding.BigEndianUnicode },
            { BigEndianUtf32, new UTF32Encoding(bigEndian: true, byteOrderMark: true) },
            { Default, Encoding.Default },
            { OEM, ClrFacade.GetOEMEncoding() },
            { String, Encoding.Unicode },
            { Unicode, Encoding.Unicode },
            { Unknown, Encoding.Unicode },
#pragma warning disable SYSLIB0001
            { Utf7, Encoding.UTF7 },
#pragma warning restore SYSLIB0001
            { Utf8, Encoding.Default },
            { Utf8Bom, Encoding.UTF8 },
            { Utf8NoBom, Encoding.Default },
            { Utf32, Encoding.UTF32 },  
        };

        /// <summary>
        /// Retrieve the encoding parameter from the command line
        /// it throws if the encoding does not match the known ones.
        /// </summary>
        /// <returns>A System.Text.Encoding object (null if no encoding specified).</returns>
        internal static Encoding Convert(Cmdlet cmdlet, string encoding)
        {
            if (string.IsNullOrEmpty(encoding))
            {
                // no parameter passed, default to UTF8
                return Encoding.Default;
            }

            if (encodingMap.TryGetValue(encoding, out Encoding foundEncoding))
            {
                WarnIfUtf7Encoding(cmdlet, foundEncoding);

                return foundEncoding;
            }

            // error condition: unknown encoding value
            string validEncodingValues = string.Join(", ", TabCompletionResults);
            string msg = StringUtil.Format(PathUtilsStrings.OutFile_WriteToFileEncodingUnknown, encoding, validEncodingValues);

            ErrorRecord errorRecord = new ErrorRecord(
                PSTraceSource.NewArgumentException("Encoding"),
                "WriteToFileEncodingUnknown",
                ErrorCategory.InvalidArgument,
                null);

            errorRecord.ErrorDetails = new ErrorDetails(msg);
            cmdlet.ThrowTerminatingError(errorRecord);

            return null;
        }

        /// <summary>
        /// Warn if the encoding is obsolete UTF-7 encoding.
        /// </summary>
        /// <param name="cmdlet">A cmdlet instance which is used to emit the warning.</param>
        /// <param name="encoding">The encoding to check for obsolescence.</param>
        internal static void WarnIfUtf7Encoding(Cmdlet cmdlet, Encoding encoding)
        {
            const int CodePageUtf7 = 65000;

            // Use the recommended pattern to check for UTF-7.
            // https://docs.microsoft.com/dotnet/core/compatibility/core-libraries/5.0/utf-7-code-paths-obsolete
            if (encoding is { CodePage: CodePageUtf7 })
            {
                cmdlet.WriteWarning(PathUtilsStrings.Utf7EncodingObsolete);
            }
        }
    }

    /// <summary>
    /// To make it easier to specify -Encoding parameter, we add an ArgumentTransformationAttribute here.
    /// When the input data is of type string and is valid to be converted to System.Text.Encoding, we do
    /// the conversion and return the converted value. Otherwise, we just return the input data.
    /// </summary>
    internal sealed class ArgumentToEncodingTransformationAttribute : ArgumentTransformationAttribute
    {
        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            switch (inputData)
            {
                case string stringName:
                    if (EncodingConversion.encodingMap.TryGetValue(stringName, out Encoding foundEncoding))
                    {
                        return foundEncoding;
                    }
                    else
                    {
                        return Encoding.GetEncoding(stringName);
                    }
                case int intName:
                    return Encoding.GetEncoding(intName);
            }

            return inputData;
        }
    }

    /// <summary>
    /// Provides the set of Encoding values for tab completion of an Encoding parameter.
    /// </summary>
    internal sealed class ArgumentEncodingCompletionsAttribute : ArgumentCompletionsAttribute
    {
        public ArgumentEncodingCompletionsAttribute() : base(
            EncodingConversion.ANSI,
            EncodingConversion.Ascii,
            EncodingConversion.BigEndianUnicode,
            EncodingConversion.BigEndianUtf32,
            EncodingConversion.OEM,
            EncodingConversion.Unicode,
            EncodingConversion.Utf7,
            EncodingConversion.Utf8,
            EncodingConversion.Utf8Bom,
            EncodingConversion.Utf8NoBom,
            EncodingConversion.Utf32
        )
        { }
    }
}
