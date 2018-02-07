

namespace System.Management.Automation.Internal
{
    using System;
    using System.Threading;
    using System.Runtime.InteropServices;
    using System.Collections.ObjectModel;
    using System.Management.Automation.Runspaces;
    using System.Management.Automation;

    /// <summary>
    /// A PipelineReader for an ObjectStream
    /// </summary>
    /// <remarks>
    /// This class is not safe for multi-threaded operations.
    /// </remarks>
    internal abstract class ObjectReaderBase<T> : PipelineReader<T>, IDisposable
    {
        /// <summary>
        /// Construct with an existing ObjectStream
        /// </summary>
        /// <param name="stream">the stream to read</param>
        /// <exception cref="ArgumentNullException">Thrown if the specified stream is null</exception>
        public ObjectReaderBase([In, Out] ObjectStreamBase stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException("stream", "stream may not be null");
            }

            _stream = stream;
        }

        #region Events

        /// <summary>
        /// Event fired when objects are added to the underlying stream
        /// </summary>
        public override event EventHandler DataReady
        {
            add
            {
                lock (_monitorObject)
                {
                    bool firstRegistrant = (null == InternalDataReady);
                    InternalDataReady += value;
                    if (firstRegistrant)
                    {
                        _stream.DataReady += new EventHandler(this.OnDataReady);
                    }
                }
            }
            remove
            {
                lock (_monitorObject)
                {
                    InternalDataReady -= value;
                    if (null == InternalDataReady)
                    {
                        _stream.DataReady -= new EventHandler(this.OnDataReady);
                    }
                }
            }
        }

        public event EventHandler InternalDataReady = null;

        #endregion Events

        #region Public Properties

        /// <summary>
        /// Waitable handle for caller's to block until data is ready to read from the underlying stream
        /// </summary>
        public override WaitHandle WaitHandle
        {
            get
            {
                return _stream.ReadHandle;
            }
        }

        /// <summary>
        /// Check if the stream is closed and contains no data.
        /// </summary>
        /// <value>True if the stream is closed and contains no data, otherwise; false.</value>
        /// <remarks>
        /// Attempting to read from the underlying stream if EndOfPipeline is true returns
        /// zero objects.
        /// </remarks>
        public override bool EndOfPipeline
        {
            get
            {
                return _stream.EndOfPipeline;
            }
        }

        /// <summary>
        /// Check if the stream is open for further writes.
        /// </summary>
        /// <value>true if the underlying stream is open, otherwise; false.</value>
        /// <remarks>
        /// The underlying stream may be readable after it is closed if data remains in the
        /// internal buffer. Check <see cref="EndOfPipeline"/> to determine if
        /// the underlying stream is closed and contains no data.
        /// </remarks>
        public override bool IsOpen
        {
            get
            {
                return _stream.IsOpen;
            }
        }

        /// <summary>
        /// Returns the number of objects in the underlying stream
        /// </summary>
        public override int Count
        {
            get
            {
                return _stream.Count;
            }
        }

        /// <summary>
        /// Get the capacity of the stream
        /// </summary>
        /// <value>
        /// The capacity of the stream.
        /// </value>
        /// <remarks>
        /// The capacity is the number of objects that stream may contain at one time.  Once this
        /// limit is reached, attempts to write into the stream block until buffer space
        /// becomes available.
        /// </remarks>
        public override int MaxCapacity
        {
            get
            {
                return _stream.MaxCapacity;
            }
        }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Close the stream
        /// </summary>
        /// <remarks>
        /// Causes subsequent calls to IsOpen to return false and calls to
        /// a write operation to throw an ObjectDisposedException.
        /// All calls to Close() after the first call are silently ignored.
        /// </remarks>
        /// <exception cref="ObjectDisposedException">
        /// The stream is already disposed
        /// </exception>
        public override void Close()
        {
            // 2003/09/02-JonN added call to close underlying stream
            _stream.Close();
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Handle DataReady events from the underlying stream
        /// </summary>
        /// <param name="sender">The stream raising the event</param>
        /// <param name="args">standard event args.</param>
        private void OnDataReady(object sender, EventArgs args)
        {
            // call any event handlers on this, replacing the
            // ObjectStream sender with 'this' since receivers
            // are expecting a PipelineReader<object>
            InternalDataReady.SafeInvoke(this, args);
        }

        #endregion Private Methods

        #region Private fields

        /// <summary>
        /// The underlying stream
        /// </summary>
        /// <remarks>Can never be null</remarks>
        protected ObjectStreamBase _stream;

        /// <summary>
        /// This object is used to acquire an exclusive lock
        /// on event handler registration.
        /// </summary>
        /// <remarks>
        /// Note that we lock _monitorObject rather than "this" so that
        /// we are protected from outside code interfering in our
        /// critical section.  Thanks to Wintellect for the hint.
        /// </remarks>
        private object _monitorObject = new Object();

        #endregion Private fields

        #region IDisposable

        /// <summary>
        /// public method for dispose
        /// </summary>
        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// release all resources
        /// </summary>
        /// <param name="disposing">if true, release all managed resources</param>
        protected abstract void Dispose(bool disposing);

        #endregion IDisposable
    } // ObjectReaderBase

    /// <summary>
    /// A PipelineReader reading objects from an ObjectStream
    /// </summary>
    /// <remarks>
    /// This class is not safe for multi-threaded operations.
    /// </remarks>
    internal class ObjectReader : ObjectReaderBase<object>
    {
        #region ctor
        /// <summary>
        /// Construct with an existing ObjectStream
        /// </summary>
        /// <param name="stream">the stream to read</param>
        /// <exception cref="ArgumentNullException">Thrown if the specified stream is null</exception>
        public ObjectReader([In, Out] ObjectStream stream)
            : base(stream)
        { }
        #endregion ctor

        /// <summary>
        /// Read at most <paramref name="count"/> objects
        /// </summary>
        /// <param name="count">The maximum number of objects to read</param>
        /// <returns>The objects read</returns>
        /// <remarks>
        /// This method blocks if the number of objects in the stream is less than <paramref name="count"/>
        /// and the stream is not closed.
        /// </remarks>
        public override Collection<object> Read(int count)
        {
            return _stream.Read(count);
        }

        /// <summary>
        /// Read a single object from the stream
        /// </summary>
        /// <returns>the next object in the stream</returns>
        /// <remarks>This method blocks if the stream is empty</remarks>
        public override object Read()
        {
            return _stream.Read();
        }

        /// <summary>
        /// Blocks until the pipeline closes and reads all objects.
        /// </summary>
        /// <returns>A collection of zero or more objects.</returns>
        /// <remarks>
        /// If the stream is empty, an empty collection is returned.
        /// </remarks>
        public override Collection<object> ReadToEnd()
        {
            return _stream.ReadToEnd();
        }

        /// <summary>
        /// Reads all objects currently in the stream, but does not block.
        /// </summary>
        /// <returns>A collection of zero or more objects.</returns>
        /// <remarks>
        /// This method performs a read of all objects currently in the
        /// stream. The method will block until exclusive access to the
        /// stream is acquired.  If there are no objects in the stream,
        /// an empty collection is returned.
        /// </remarks>
        public override Collection<object> NonBlockingRead()
        {
            return _stream.NonBlockingRead(Int32.MaxValue);
        }

        /// <summary>
        /// Reads objects currently in the stream, but does not block.
        /// </summary>
        /// <returns>A collection of zero or more objects.</returns>
        /// <remarks>
        /// This method performs a read of objects currently in the
        /// stream. The method will block until exclusive access to the
        /// stream is acquired.  If there are no objects in the stream,
        /// an empty collection is returned.
        /// </remarks>
        /// <param name="maxRequested">
        /// Return no more than maxRequested objects.
        /// </param>
        public override Collection<object> NonBlockingRead(int maxRequested)
        {
            return _stream.NonBlockingRead(maxRequested);
        }

        /// <summary>
        /// Peek the next object
        /// </summary>
        /// <returns>The next object in the stream or ObjectStream.EmptyObject if the stream is empty</returns>
        public override object Peek()
        {
            return _stream.Peek();
        }

        /// <summary>
        /// release all resources
        /// </summary>
        /// <param name="disposing">if true, release all managed resources</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _stream.Close();
            }
        }
    } // ObjectReader

    /// <summary>
    /// A PipelineReader reading PSObjects from an ObjectStream
    /// </summary>
    /// <remarks>
    /// This class is not safe for multi-threaded operations.
    /// </remarks>
    internal class PSObjectReader : ObjectReaderBase<PSObject>
    {
        #region ctor
        /// <summary>
        /// Construct with an existing ObjectStream
        /// </summary>
        /// <param name="stream">the stream to read</param>
        /// <exception cref="ArgumentNullException">Thrown if the specified stream is null</exception>
        public PSObjectReader([In, Out] ObjectStream stream)
            : base(stream)
        { }
        #endregion ctor

        /// <summary>
        /// Read at most <paramref name="count"/> objects
        /// </summary>
        /// <param name="count">The maximum number of objects to read</param>
        /// <returns>The objects read</returns>
        /// <remarks>
        /// This method blocks if the number of objects in the stream is less than <paramref name="count"/>
        /// and the stream is not closed.
        /// </remarks>
        public override Collection<PSObject> Read(int count)
        {
            return MakePSObjectCollection(_stream.Read(count));
        }

        /// <summary>
        /// Read a single PSObject from the stream
        /// </summary>
        /// <returns>the next PSObject in the stream</returns>
        /// <remarks>This method blocks if the stream is empty</remarks>
        public override PSObject Read()
        {
            return MakePSObject(_stream.Read());
        }

        /// <summary>
        /// Blocks until the pipeline closes and reads all objects.
        /// </summary>
        /// <returns>A collection of zero or more objects.</returns>
        /// <remarks>
        /// If the stream is empty, an empty collection is returned.
        /// </remarks>
        public override Collection<PSObject> ReadToEnd()
        {
            return MakePSObjectCollection(_stream.ReadToEnd());
        }

        /// <summary>
        /// Reads all objects currently in the stream, but does not block.
        /// </summary>
        /// <returns>A collection of zero or more objects.</returns>
        /// <remarks>
        /// This method performs a read of all objects currently in the
        /// stream. The method will block until exclusive access to the
        /// stream is acquired.  If there are no objects in the stream,
        /// an empty collection is returned.
        /// </remarks>
        public override Collection<PSObject> NonBlockingRead()
        {
            return MakePSObjectCollection(_stream.NonBlockingRead(Int32.MaxValue));
        }

        /// <summary>
        /// Reads objects currently in the stream, but does not block.
        /// </summary>
        /// <returns>A collection of zero or more objects.</returns>
        /// <remarks>
        /// This method performs a read of objects currently in the
        /// stream. The method will block until exclusive access to the
        /// stream is acquired.  If there are no objects in the stream,
        /// an empty collection is returned.
        /// </remarks>
        /// <param name="maxRequested">
        /// Return no more than maxRequested objects.
        /// </param>
        public override Collection<PSObject> NonBlockingRead(int maxRequested)
        {
            return MakePSObjectCollection(_stream.NonBlockingRead(maxRequested));
        }

        /// <summary>
        /// Peek the next PSObject
        /// </summary>
        /// <returns>The next PSObject in the stream or ObjectStream.EmptyObject if the stream is empty</returns>
        public override PSObject Peek()
        {
            return MakePSObject(_stream.Peek());
        }

        /// <summary>
        /// release all resources
        /// </summary>
        /// <param name="disposing">if true, release all managed resources</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _stream.Close();
            }
        }

        #region Private
        private static PSObject MakePSObject(object o)
        {
            if (null == o)
                return null;

            return PSObject.AsPSObject(o);
        }

        // It might ultimately be more efficient to
        // make ObjectStream generic and convert the objects to PSObject
        // before inserting them into the initial Collection, so that we
        // don't have to convert the collection later.
        private static Collection<PSObject> MakePSObjectCollection(
            Collection<object> coll)
        {
            if (null == coll)
                return null;
            Collection<PSObject> retval = new Collection<PSObject>();
            foreach (object o in coll)
            {
                retval.Add(MakePSObject(o));
            }
            return retval;
        }
        #endregion Private
    } // PSObjectReader

    /// <summary>
    /// A ObjectReader for a PSDataCollection ObjectStream
    /// </summary>
    /// <remarks>
    /// PSDataCollection is introduced after 1.0. PSDataCollection is
    /// used to store data which can be used with different
    /// commands concurrently.
    /// Only Read() operation is supported currently.
    /// </remarks>
    internal class PSDataCollectionReader<DataStoreType, ReturnType>
        : ObjectReaderBase<ReturnType>
    {
        #region Private Data

        private PSDataCollectionEnumerator<DataStoreType> _enumerator;

        #endregion

        #region ctor
        /// <summary>
        /// Construct with an existing ObjectStream
        /// </summary>
        /// <param name="stream">the stream to read</param>
        /// <exception cref="ArgumentNullException">Thrown if the specified stream is null</exception>
        public PSDataCollectionReader(PSDataCollectionStream<DataStoreType> stream)
            : base(stream)
        {
            System.Management.Automation.Diagnostics.Assert(null != stream.ObjectStore,
                "Stream should have a valid data store");
            _enumerator = (PSDataCollectionEnumerator<DataStoreType>)stream.ObjectStore.GetEnumerator();
        }

        #endregion ctor

        /// <summary>
        /// This method is not supported.
        /// </summary>
        /// <param name="count">The maximum number of objects to read</param>
        /// <returns>The objects read</returns>
        public override Collection<ReturnType> Read(int count)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Read a single object from the stream.
        /// </summary>
        /// <returns>
        /// The next object in the buffer or AutomationNull if buffer is closed
        /// and data is not available.
        /// </returns>
        /// <remarks>
        /// This method blocks if the buffer is empty.
        /// </remarks>
        public override ReturnType Read()
        {
            object result = AutomationNull.Value;
            if (_enumerator.MoveNext())
            {
                result = _enumerator.Current;
            }

            return ConvertToReturnType(result);
        }

        /// <summary>
        /// This method is not supported.
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public override Collection<ReturnType> ReadToEnd()
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// This method is not supported.
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public override Collection<ReturnType> NonBlockingRead()
        {
            return NonBlockingRead(Int32.MaxValue);
        }

        /// <summary>
        /// This method is not supported.
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        /// <param name="maxRequested">
        /// Return no more than maxRequested objects.
        /// </param>
        public override Collection<ReturnType> NonBlockingRead(int maxRequested)
        {
            if (maxRequested < 0)
            {
                throw PSTraceSource.NewArgumentOutOfRangeException("maxRequested", maxRequested);
            }

            if (maxRequested == 0)
            {
                return new Collection<ReturnType>();
            }
            Collection<ReturnType> results = new Collection<ReturnType>();
            int readCount = maxRequested;

            while (readCount > 0)
            {
                if (_enumerator.MoveNext(false))
                {
                    results.Add(ConvertToReturnType(_enumerator.Current));
                    continue;
                }

                break;
            }

            return results;
        }

        /// <summary>
        /// This method is not supported.
        /// </summary>
        /// <returns></returns>
        public override ReturnType Peek()
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// release all resources
        /// </summary>
        /// <param name="disposing">if true, release all managed resources</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _stream.Close();
            }
        }

        private ReturnType ConvertToReturnType(object inputObject)
        {
            Type resultType = typeof(ReturnType);
            if (typeof(PSObject) == resultType || typeof(object) == resultType)
            {
                ReturnType result;
                LanguagePrimitives.TryConvertTo(inputObject, out result);
                return result;
            }

            System.Management.Automation.Diagnostics.Assert(false,
                "ReturnType should be either object or PSObject only");
            throw PSTraceSource.NewNotSupportedException();
        }
    }

    /// <summary>
    /// A ObjectReader for a PSDataCollection ObjectStream
    /// </summary>
    /// <remarks>
    /// PSDataCollection is introduced after 1.0. PSDataCollection is
    /// used to store data which can be used with different
    /// commands concurrently.
    /// Only Read() operation is supported currently.
    /// </remarks>
    internal class PSDataCollectionPipelineReader<DataStoreType, ReturnType>
        : ObjectReaderBase<ReturnType>
    {
        #region Private Data

        private PSDataCollection<DataStoreType> _datastore;

        #endregion Private Data

        #region ctor
        /// <summary>
        /// Construct with an existing ObjectStream
        /// </summary>
        /// <param name="stream">the stream to read</param>
        /// <param name="computerName"></param>
        /// <param name="runspaceId"></param>
        internal PSDataCollectionPipelineReader(PSDataCollectionStream<DataStoreType> stream,
            String computerName, Guid runspaceId)
            : base(stream)
        {
            System.Management.Automation.Diagnostics.Assert(null != stream.ObjectStore,
                "Stream should have a valid data store");
            _datastore = stream.ObjectStore;
            ComputerName = computerName;
            RunspaceId = runspaceId;
        }

        #endregion ctor

        /// <summary>
        /// Computer name passed in by the pipeline which
        /// created this reader
        /// </summary>
        internal String ComputerName { get; }

        /// <summary>
        /// Runspace Id passed in by the pipeline which
        /// created this reader
        /// </summary>
        internal Guid RunspaceId { get; }

        /// <summary>
        /// This method is not supported.
        /// </summary>
        /// <param name="count">The maximum number of objects to read</param>
        /// <returns>The objects read</returns>
        public override Collection<ReturnType> Read(int count)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Read a single object from the stream.
        /// </summary>
        /// <returns>
        /// The next object in the buffer or AutomationNull if buffer is closed
        /// and data is not available.
        /// </returns>
        /// <remarks>
        /// This method blocks if the buffer is empty.
        /// </remarks>
        public override ReturnType Read()
        {
            object result = AutomationNull.Value;
            if (_datastore.Count > 0)
            {
                Collection<DataStoreType> resultCollection = _datastore.ReadAndRemove(1);

                // ReadAndRemove returns a Collection<DataStoreType> type but we
                // just want the single object contained in the collection.
                if (resultCollection.Count == 1)
                {
                    result = resultCollection[0];
                }
            }

            return ConvertToReturnType(result);
        }

        /// <summary>
        /// This method is not supported.
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public override Collection<ReturnType> ReadToEnd()
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// This method is not supported.
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public override Collection<ReturnType> NonBlockingRead()
        {
            return NonBlockingRead(Int32.MaxValue);
        }

        /// <summary>
        /// This method is not supported.
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        /// <param name="maxRequested">
        /// Return no more than maxRequested objects.
        /// </param>
        public override Collection<ReturnType> NonBlockingRead(int maxRequested)
        {
            if (maxRequested < 0)
            {
                throw PSTraceSource.NewArgumentOutOfRangeException("maxRequested", maxRequested);
            }

            if (maxRequested == 0)
            {
                return new Collection<ReturnType>();
            }
            Collection<ReturnType> results = new Collection<ReturnType>();
            int readCount = maxRequested;

            while (readCount > 0)
            {
                if (_datastore.Count > 0)
                {
                    results.Add(ConvertToReturnType((_datastore.ReadAndRemove(1))[0]));
                    readCount--;
                    continue;
                }

                break;
            }

            return results;
        }

        /// <summary>
        /// This method is not supported.
        /// </summary>
        /// <returns></returns>
        public override ReturnType Peek()
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Converts to the return type based on language primitives
        /// </summary>
        /// <param name="inputObject">input object to convert</param>
        /// <returns>input object converted to the specified return type</returns>
        private ReturnType ConvertToReturnType(object inputObject)
        {
            Type resultType = typeof(ReturnType);
            if (typeof(PSObject) == resultType || typeof(object) == resultType)
            {
                ReturnType result;
                LanguagePrimitives.TryConvertTo(inputObject, out result);
                return result;
            }

            System.Management.Automation.Diagnostics.Assert(false,
                "ReturnType should be either object or PSObject only");
            throw PSTraceSource.NewNotSupportedException();
        }

        #region IDisposable

        /// <summary>
        /// release all resources
        /// </summary>
        /// <param name="disposing">if true, release all managed resources</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _datastore.Dispose();
            }
        }

        #endregion IDisposable
    }
}