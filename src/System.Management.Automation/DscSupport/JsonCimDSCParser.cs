// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Security;

namespace Microsoft.PowerShell.DesiredStateConfiguration
{
    /// <summary>
    /// Class that does high level Cim schema parsing
    /// </summary>
    internal class CimDSCParser
    {
        private JsonDeserializer _json_deserializer;
        
        internal CimDSCParser()
        {
            _json_deserializer = JsonDeserializer.Create();
        }

        internal List<PSObject> ParseSchemaJson(string filePath)
        {
            string json = System.IO.File.ReadAllText(filePath);
            try
            {
                string fileNameDefiningClass = Path.GetFileNameWithoutExtension(filePath);
                int dotIndex = fileNameDefiningClass.IndexOf('.');
                if (dotIndex != -1)
                {
                    fileNameDefiningClass = fileNameDefiningClass.Substring(0, dotIndex);
                }

                var result = new List<PSObject>(_json_deserializer.DeserializeClasses(json));
                foreach (dynamic c in result)
                {
                    string superClassName = c.CimSuperClassName;
                    string className = c.CimSystemProperties.ClassName;
                    if ((superClassName != null) && (superClassName.Equals("OMI_BaseResource", StringComparison.OrdinalIgnoreCase)))
                    {
                        // Get the name of the file without schema.mof/json extension
                        if (!(className.Equals(fileNameDefiningClass, StringComparison.OrdinalIgnoreCase)))
                        {
                            PSInvalidOperationException e = PSTraceSource.NewInvalidOperationException(
                                ParserStrings.ClassNameNotSameAsDefiningFile, className, fileNameDefiningClass);
                            throw e;
                        }
                    }
                }

                return result;
            }
            catch (Exception exception)
            {
                PSInvalidOperationException e = PSTraceSource.NewInvalidOperationException(
                    exception, ParserStrings.CimDeserializationError, filePath);

                e.SetErrorId("CimDeserializationError");
                throw e;
            }
        }

        /// <summary>
        /// Make sure that the instance conforms to the the schema.
        /// </summary>
        /// <param name="classText"></param>
        internal void ValidateInstanceText(string classText)
        {
            throw new NotImplementedException("Instance parsing/validation is not yet suported by JSON-based parser");
        }
    }
}
