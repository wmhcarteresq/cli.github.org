// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Reflection;

namespace Microsoft.PowerShell.Commands
{
    /// <summary>
    /// The ConvertFrom-Json command.
    /// This command convert a Json string representation to a JsonObject.
    /// </summary>
    [Cmdlet(VerbsData.ConvertFrom, "Json", HelpUri = "https://go.microsoft.com/fwlink/?LinkID=217031", RemotingCapability = RemotingCapability.None)]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly")]
    public class ConvertFromJsonCommand : Cmdlet
    {
        #region parameters

        /// <summary>
        /// Gets or sets the InputString property.
        /// </summary>
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
        [AllowEmptyString]
        public string InputObject { get; set; }

        /// <summary>
        /// InputObjectBuffer buffers all InputObject contents available in the pipeline.
        /// </summary>
        private List<string> _inputObjectBuffer = new List<string>();

        /// <summary>
        /// Returned data structure is a Hashtable instead a CustomPSObject.
        /// </summary>
        [Parameter()]
        public SwitchParameter AsHashtable { get; set; }

        #endregion parameters

        #region overrides

        /// <summary>
        ///  Buffers InputObjet contents available in the pipeline.
        /// </summary>
        protected override void ProcessRecord()
        {
            _inputObjectBuffer.Add(InputObject);
        }

        /// <summary>
        /// The main execution method for the ConvertFrom-Json command.
        /// </summary>
        protected override void EndProcessing()
        {
            // When Input is provided through pipeline, the input can be represented in the following two ways:
            // 1. Each input in the collection is a complete Json content. There can be multiple inputs of this format.
            // 2. The complete input is a collection which represents a single Json content. This is typically the majority of the case.
            if (_inputObjectBuffer.Count > 0)
            {
                if (_inputObjectBuffer.Count == 1)
                {
                    ConvertFromJsonHelper(_inputObjectBuffer[0]);
                }
                else
                {
                    bool successfullyConverted = false;
                    try
                    {
                        // Try to deserialize the first element.
                        successfullyConverted = ConvertFromJsonHelper(_inputObjectBuffer[0]);
                    }
                    catch (ArgumentException)
                    {
                        // The first input string does not represent a complete Json Syntax.
                        // Hence consider the the entire input as a single Json content.
                    }

                    if (successfullyConverted)
                    {
                        for (int index = 1; index < _inputObjectBuffer.Count; index++)
                        {
                            ConvertFromJsonHelper(_inputObjectBuffer[index]);
                        }
                    }
                    else
                    {
                        // Process the entire input as a single Json content.
                        ConvertFromJsonHelper(string.Join(System.Environment.NewLine, _inputObjectBuffer.ToArray()));
                    }
                }
            }
        }

        /// <summary>
        /// ConvertFromJsonHelper is a helper method to convert to Json input to .Net Type.
        /// </summary>
        /// <param name="input">Input String.</param>
        /// <returns>True if successfully converted, else returns false.</returns>
        private bool ConvertFromJsonHelper(string input)
        {
            ErrorRecord error = null;
            object result = JsonObject.ConvertFromJson(input, AsHashtable.IsPresent, out error);

            if (error != null)
            {
                ThrowTerminatingError(error);
            }

            WriteObject(result);
            return (result != null);
        }

        #endregion overrides
    }
}
