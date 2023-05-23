// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Runtime.Serialization;

namespace System.Management.Automation
{
    /// <summary>
    /// This is a wrapper for exception class
    /// <see cref="System.InvalidOperationException"/>
    /// which provides additional information via
    /// <see cref="System.Management.Automation.IContainsErrorRecord"/>.
    /// </summary>
    /// <remarks>
    /// Instances of this exception class are usually generated by the
    /// PowerShell Engine.  It is unusual for code outside the PowerShell Engine
    /// to create an instance of this class.
    /// </remarks>
    public class PSInvalidOperationException
            : InvalidOperationException, IContainsErrorRecord
    {
        #region ctor
        /// <summary>
        /// Initializes a new instance of the PSInvalidOperationException class.
        /// </summary>
        /// <returns>Constructed object.</returns>
        public PSInvalidOperationException()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the PSInvalidOperationException class.
        /// </summary>
        /// <param name="message"></param>
        /// <returns>Constructed object.</returns>
        public PSInvalidOperationException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PSInvalidOperationException class.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        /// <returns>Constructed object.</returns>
        public PSInvalidOperationException(string message,
                                            Exception innerException)
                : base(message, innerException)
        {
        }

        /// <summary>
        /// Initializes a new instance of the PSInvalidOperationException class.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        /// <param name="errorId"></param>
        /// <param name="errorCategory"></param>
        /// <param name="target"></param>
        /// <returns>Constructed object.</returns>
        internal PSInvalidOperationException(string message, Exception innerException, string errorId, ErrorCategory errorCategory, object target)
            : base(message, innerException)
        {
            _errorId = errorId;
            _errorCategory = errorCategory;
            _target = target;
        }
        #endregion ctor

        /// <summary>
        /// Additional information about the error.
        /// </summary>
        /// <value></value>
        /// <remarks>
        /// Note that ErrorRecord.Exception is
        /// <see cref="System.Management.Automation.ParentContainsErrorRecordException"/>.
        /// </remarks>
        public ErrorRecord ErrorRecord
        {
            get
            {
                _errorRecord ??= new ErrorRecord(
                    new ParentContainsErrorRecordException(this),
                    _errorId,
                    _errorCategory,
                    _target);

                return _errorRecord;
            }
        }

        private ErrorRecord _errorRecord;
        private string _errorId = "InvalidOperation";

        internal void SetErrorId(string errorId)
        {
            _errorId = errorId;
        }

        private readonly ErrorCategory _errorCategory = ErrorCategory.InvalidOperation;
        private readonly object _target = null;
    }
}
