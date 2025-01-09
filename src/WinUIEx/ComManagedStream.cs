// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.

using System;
using System.Buffers;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Windows.Win32.Foundation;
using Windows.Win32.System.Com;

namespace WinUIEx
{
    internal sealed unsafe class ComManagedStream : IStream.Interface
    {
        public static Guid IStreamInterfaceGUID = new("0000000C-0000-0000-C000-000000000046");

        private readonly Stream _dataStream;

        // To support seeking ahead of the stream length
        private long _virtualPosition = -1;

        internal ComManagedStream(Stream stream, bool makeSeekable = false)
        {
            if (makeSeekable && !stream.CanSeek)
            {
                // Copy to a memory stream so we can seek
                MemoryStream memoryStream = new();
                stream.CopyTo(memoryStream);
                _dataStream = memoryStream;
            }
            else
            {
                _dataStream = stream;
            }
        }

        private void ActualizeVirtualPosition()
        {
            if (_virtualPosition == -1)
                return;

            if (_virtualPosition > _dataStream.Length)
                _dataStream.SetLength(_virtualPosition);

            _dataStream.Position = _virtualPosition;

            _virtualPosition = -1;
        }
        public Stream GetDataStream() => _dataStream;

        HRESULT IStream.Interface.Clone(IStream** ppstm)
        {
            if (ppstm is null)
            {
                return HRESULT.E_POINTER;
            }

            // The cloned object should have the same current "position"
            *ppstm = ComHelpers.GetComPointer<IStream>(
                new ComManagedStream(_dataStream) { _virtualPosition = _virtualPosition });

            return HRESULT.S_OK;
        }

        HRESULT IStream.Interface.Commit(uint grfCommitFlags)
        {
            _dataStream.Flush();

            // Extend the length of the file if needed.
            ActualizeVirtualPosition();
            return HRESULT.S_OK;
        }

        HRESULT IStream.Interface.CopyTo(IStream* pstm, ulong cb, ulong* pcbRead, ulong* pcbWritten)
        {
            if (pstm is null)
            {
                return HRESULT.STG_E_INVALIDPOINTER;
            }

            using BufferScope<byte> buffer = new(4096);

            ulong remaining = cb;
            ulong totalWritten = 0;
            ulong totalRead = 0;

            fixed (byte* b = buffer)
            {
                while (remaining > 0)
                {
                    uint read = remaining < (ulong)buffer.Length ? (uint)remaining : (uint)buffer.Length;

                    ((IStream.Interface)this).Read(b, read, &read);
                    remaining -= read;
                    totalRead += read;

                    if (read == 0)
                    {
                        break;
                    }

                    uint written;
                    pstm->Write(b, read, &written).ThrowOnFailure();
                    totalWritten += written;
                }
            }

            if (pcbRead is not null)
                *pcbRead = totalRead;

            if (pcbWritten is not null)
                *pcbWritten = totalWritten;

            return HRESULT.S_OK;
        }

        HRESULT ISequentialStream.Interface.Read(void* pv, uint cb, uint* pcbRead)
        {
            if (pv is null)
            {
                return HRESULT.STG_E_INVALIDPOINTER;
            }

            ActualizeVirtualPosition();

            Span<byte> buffer = new(pv, checked((int)cb));
            int read = _dataStream.Read(buffer);

            if (pcbRead is not null)
                *pcbRead = (uint)read;
            return HRESULT.S_OK;
        }

        HRESULT IStream.Interface.Read(void* pv, uint cb, uint* pcbRead)
            => ((ISequentialStream.Interface)this).Read(pv, cb, pcbRead);

        HRESULT IStream.Interface.Seek(long dlibMove, SeekOrigin dwOrigin, ulong* plibNewPosition)
        {
            long position = _virtualPosition == -1 ? _dataStream.Position : _virtualPosition;
            long length = _dataStream.Length;

            switch (dwOrigin)
            {
                case SeekOrigin.Begin:
                    if (dlibMove <= length)
                    {
                        _dataStream.Position = dlibMove;
                        _virtualPosition = -1;
                    }
                    else
                    {
                        _virtualPosition = dlibMove;
                    }

                    break;
                case SeekOrigin.End:
                    if (dlibMove <= 0)
                    {
                        _dataStream.Position = length + dlibMove;
                        _virtualPosition = -1;
                    }
                    else
                    {
                        _virtualPosition = length + dlibMove;
                    }

                    break;
                case SeekOrigin.Current:
                    if (dlibMove + position <= length)
                    {
                        _dataStream.Position = position + dlibMove;
                        _virtualPosition = -1;
                    }
                    else
                    {
                        _virtualPosition = dlibMove + position;
                    }

                    break;
            }

            if (plibNewPosition is null)
                return HRESULT.S_OK;

            *plibNewPosition = _virtualPosition == -1 ? (ulong)_dataStream.Position : (ulong)_virtualPosition;
            return HRESULT.S_OK;
        }

        HRESULT IStream.Interface.SetSize(ulong libNewSize)
        {
            _dataStream.SetLength(checked((long)libNewSize));
            return HRESULT.S_OK;
        }

        HRESULT IStream.Interface.Stat(STATSTG* pstatstg, uint grfStatFlag)
        {
            if (pstatstg is null)
            {
                return HRESULT.STG_E_INVALIDPOINTER;
            }

            *pstatstg = new STATSTG
            {
                cbSize = (ulong)_dataStream.Length,
                type = (uint)STGTY.STGTY_STREAM,

                // Default read/write access is READ, which == 0
                grfMode = _dataStream.CanWrite
                    ? _dataStream.CanRead
                        ? STGM.STGM_READWRITE
                        : STGM.STGM_WRITE
                    : STGM.STGM_READ
            };

            if ((STATFLAG)grfStatFlag == STATFLAG.STATFLAG_DEFAULT)
            {
                // Caller wants a name
                pstatstg->pwcsName = (char*)Marshal.StringToCoTaskMemUni(_dataStream is FileStream fs ? fs.Name : _dataStream.ToString());
            }

            return HRESULT.S_OK;
        }

        /// Returns HRESULT.STG_E_INVALIDFUNCTION as a documented way to say we don't support locking
        HRESULT IStream.Interface.LockRegion(ulong libOffset, ulong cb, uint dwLockType) => HRESULT.STG_E_INVALIDFUNCTION;

        // We never report ourselves as Transacted, so we can just ignore this.
        HRESULT IStream.Interface.Revert() => HRESULT.S_OK;

        /// Returns HRESULT.STG_E_INVALIDFUNCTION as a documented way to say we don't support locking
        HRESULT IStream.Interface.UnlockRegion(ulong libOffset, ulong cb, uint dwLockType) => HRESULT.STG_E_INVALIDFUNCTION;

        HRESULT ISequentialStream.Interface.Write(void* pv, uint cb, uint* pcbWritten)
        {
            if (pv is null)
            {
                return HRESULT.STG_E_INVALIDPOINTER;
            }

            ActualizeVirtualPosition();

            ReadOnlySpan<byte> buffer = new(pv, checked((int)cb));
            _dataStream.Write(buffer);

            if (pcbWritten is not null)
                *pcbWritten = cb;
            return HRESULT.S_OK;
        }

        HRESULT IStream.Interface.Write(void* pv, uint cb, uint* pcbWritten)
            => ((ISequentialStream.Interface)this).Write(pv, cb, pcbWritten);
    }

    internal ref struct BufferScope<T>
    {
        private T[]? _array;
        private Span<T> _span;

        public BufferScope(int minimumLength)
        {
            _array = ArrayPool<T>.Shared.Rent(minimumLength);
            _span = _array;
        }

        /// <summary>
        ///  Create the <see cref="BufferScope{T}"/> with an initial buffer. Useful for creating with an initial stack
        ///  allocated buffer.
        /// </summary>
        public BufferScope(Span<T> initialBuffer)
        {
            _array = null;
            _span = initialBuffer;
        }

        /// <summary>
        ///  Create the <see cref="BufferScope{T}"/> with an initial buffer. Useful for creating with an initial stack
        ///  allocated buffer.
        /// </summary>
        /// <remarks>
        ///  <para>
        ///   <example>
        ///    <para>Creating with a stack allocated buffer:</para>
        ///    <code>using BufferScope&lt;char> buffer = new(stackalloc char[64]);</code>
        ///   </example>
        ///  </para>
        ///  <para>
        ///   Stack allocated buffers should be kept small to avoid overflowing the stack.
        ///  </para>
        /// </remarks>
        /// <param name="initialBuffer">The initial buffer to use.</param>
        /// <param name="minimumLength">
        ///  The required minimum length. If the <paramref name="initialBuffer"/> is not large enough, this will rent from
        ///  the shared <see cref="ArrayPool{T}"/>.
        /// </param>
        public BufferScope(Span<T> initialBuffer, int minimumLength)
        {
            if (initialBuffer.Length >= minimumLength)
            {
                _array = null;
                _span = initialBuffer;
            }
            else
            {
                _array = ArrayPool<T>.Shared.Rent(minimumLength);
                _span = _array;
            }
        }

        /// <summary>
        ///  Ensure that the buffer has enough space for <paramref name="capacity"/> number of elements.
        /// </summary>
        /// <remarks>
        ///  <para>
        ///   Consider if creating new <see cref="BufferScope{T}"/> instances is possible and cleaner than using
        ///   this method.
        ///  </para>
        /// </remarks>
        /// <param name="capacity">The minimum capacity required.</param>
        /// <param name="copy">True to copy the existing elements when new space is allocated.</param>
        public unsafe void EnsureCapacity(int capacity, bool copy = false)
        {
            if (_span!.Length >= capacity)
            {
                return;
            }

            T[] newArray = ArrayPool<T>.Shared.Rent(capacity);
            if (copy)
            {
                _span.CopyTo(newArray);
            }

            if (_array is not null)
            {
                ArrayPool<T>.Shared.Return(_array);
            }

            _array = newArray;
            _span = _array;
        }

        public ref T this[int i] => ref _span[i];

        public readonly Span<T> this[Range range] => _span[range];

        public readonly Span<T> Slice(int start, int length) => _span.Slice(start, length);

        public readonly ref T GetPinnableReference() => ref MemoryMarshal.GetReference(_span);

        public readonly int Length => _span.Length;

        public readonly Span<T> AsSpan() => _span;

        public static implicit operator Span<T>(BufferScope<T> scope) => scope._span;

        public static implicit operator ReadOnlySpan<T>(BufferScope<T> scope) => scope._span;

        public void Dispose()
        {
            if (_array is not null)
            {
                ArrayPool<T>.Shared.Return(_array);
            }

            _array = default;
        }

        public override readonly string ToString() => _span.ToString();
    }

    internal unsafe static class ComHelpers
    {
        internal static ComScope<T> GetComScope<T>(object? @object) where T : unmanaged =>
           new(GetComPointer<T>(@object));

        /// <summary>
        ///  Gets the specified <typeparamref name="T"/> interface for the given <paramref name="object"/>. Throws if
        ///  the desired pointer can not be obtained.
        /// </summary>
        internal static T* GetComPointer<T>(object? @object) where T : unmanaged
        {
            T* result = TryGetComPointer<T>(@object, out HRESULT hr);
            hr.ThrowOnFailure();
            return result;
        }

        /// <summary>
        ///  Attempts to get the specified <typeparamref name="T"/> interface for the given <paramref name="object"/>.
        /// </summary>
        /// <returns>The requested pointer or <see langword="null"/> if unsuccessful.</returns>
        internal static T* TryGetComPointer<T>(object? @object) where T : unmanaged =>
            TryGetComPointer<T>(@object, out _);

        /// <summary>
        ///  Attempts to get the specified <typeparamref name="T"/> interface for the given <paramref name="object"/>.
        /// </summary>
        /// <param name="object">The object to get the pointer from.</param>
        /// <param name="result">
        ///  Typically either <see cref="HRESULT.S_OK"/> or <see cref="HRESULT.E_POINTER"/>. Check for success, not
        ///  specific results.
        /// </param>
        /// <returns>The requested pointer or <see langword="null"/> if unsuccessful.</returns>
        internal static T* TryGetComPointer<T>(object? @object, out HRESULT result) where T : unmanaged
        {
            if (@object is null)
            {
                result = HRESULT.E_POINTER;
                return null;
            }

            IUnknown* ccw = null;
            try
            {
                ccw = (IUnknown*)Marshal.GetIUnknownForObject(@object);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Did not find IUnknown for {@object.GetType().Name}. {ex.Message}");
            }

            if (ccw is null)
            {
                result = HRESULT.E_NOINTERFACE;
                return null;
            }

            // Now query out the requested interface
            result = ccw->QueryInterface(ComManagedStream.IStreamInterfaceGUID, out void* ppvObject);
            ccw->Release();
            return (T*)ppvObject;
        }
    }

    internal readonly unsafe ref struct ComScope<T> where T : unmanaged
    {
        /// <summary>
        ///  Gets a pointer to the IID <see cref="Guid"/> for the given <typeparamref name="T"/>.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static Guid* Get<TIn>() where TIn : unmanaged
            => (Guid*)Unsafe.AsPointer(ref Unsafe.AsRef(ComManagedStream.IStreamInterfaceGUID));

        // Keeping internal as nint allows us to use Unsafe methods to get significantly better generated code.
        private readonly nint _value;
        public T* Value => (T*)_value;
        public IUnknown* AsUnknown => (IUnknown*)_value;

        public ComScope(T* value) => _value = (nint)value;

        // Can't add an operator for IUnknown* as we have ComScope<IUnknown>.

        public static implicit operator T*(in ComScope<T> scope) => (T*)scope._value;

        public static implicit operator void*(in ComScope<T> scope) => (void*)scope._value;

        public static implicit operator nint(in ComScope<T> scope) => scope._value;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static implicit operator T**(in ComScope<T> scope) => (T**)Unsafe.AsPointer(ref Unsafe.AsRef(in scope._value));

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static implicit operator void**(in ComScope<T> scope) => (void**)Unsafe.AsPointer(ref Unsafe.AsRef(in scope._value));

        public bool IsNull => _value == 0;

        /// <summary>
        ///  Tries querying the requested interface into a new <see cref="ComScope{T}"/>.
        /// </summary>
        /// <param name="hr">The result of the query.</param>
        public ComScope<TTo> TryQuery<TTo>(out HRESULT hr) where TTo : unmanaged
        {
            ComScope<TTo> scope = new(null);
            hr = ((IUnknown*)Value)->QueryInterface(Get<TTo>(), scope);
            return scope;
        }

        /// <summary>
        ///  Queries the requested interface into a new <see cref="ComScope{T}"/>.
        /// </summary>
        public ComScope<TTo> Query<TTo>() where TTo : unmanaged
        {
            ComScope<TTo> scope = new(null);
            ((IUnknown*)Value)->QueryInterface(Get<TTo>(), scope).ThrowOnFailure();
            return scope;
        }

        /// <summary>
        ///  Attempt to create a <see cref="ComScope{T}"/> from the given COM interface.
        /// </summary>
        public static ComScope<T> TryQueryFrom<TFrom>(TFrom* from, out HRESULT hr) where TFrom : unmanaged
        {
            ComScope<T> scope = new(null);
            hr = from is null ? HRESULT.E_POINTER : ((IUnknown*)from)->QueryInterface(Get<T>(), scope);
            return scope;
        }

        /// <summary>
        ///  Create a <see cref="ComScope{T}"/> from the given COM interface. Throws on failure.
        /// </summary>
        public static ComScope<T> QueryFrom<TFrom>(TFrom* from) where TFrom : unmanaged
        {
            if (from is null)
            {
                HRESULT.E_POINTER.ThrowOnFailure();
            }

            ComScope<T> scope = new(null);
            ((IUnknown*)from)->QueryInterface(Get<T>(), scope).ThrowOnFailure();
            return scope;
        }

        /// <summary>
        ///  Simple helper for checking if a given interface is supported. Only use this if you don't intend to
        ///  use the interface, otherwise use <see cref="TryQuery{TTo}(out HRESULT)"/>.
        /// </summary>
        public bool SupportsInterface<TInterface>() where TInterface : unmanaged
        {
            if (typeof(TInterface) == typeof(T))
            {
                return true;
            }

            IUnknown* unknown;
            HRESULT hr = AsUnknown->QueryInterface(Get<TInterface>(), (void**)&unknown);

            if (hr.Succeeded)
            {
                unknown->Release();
                return true;
            }

            return false;
        }

        public void Dispose()
        {
            IUnknown* unknown = (IUnknown*)_value;

            // Really want this to be null after disposal to avoid double releases, but we also want
            // to maintain the readonly state of the struct to allow passing as `in` without creating implicit
            // copies (which would break the T** and void** operators).
            *(void**)this = null;
            if (unknown is not null)
            {
                unknown->Release();
            }
        }
    }
}
