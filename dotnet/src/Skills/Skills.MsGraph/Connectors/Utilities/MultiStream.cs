// Copyright (c) Microsoft. All rights reserved.

using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Microsoft.SemanticKernel.Skills.MsGraph.Connectors.Utilities;

public class MultiStream : Stream
{
    private long _position;
    private Queue<Stream> _streams;

    public MultiStream(IEnumerable<Stream> streams)
    {
       this._position = 0;
        this._streams = new Queue<Stream>(streams);
    }

    public override bool CanRead
    {
        get
        {
            return this._streams.Count == 0 || this._streams.Any(s => s.CanRead);
        }
    }

    public override bool CanSeek => false;

    public override bool CanWrite => false;

    public override long Length
    {
        get
        {
            long length = 0;

            foreach (Stream s in this._streams)
            {
                length += s.Length;
            }

            return length;
        }
    }

    public override long Position
    {
        get => this._position;
        set => throw new System.NotImplementedException();
    }

    public override void Flush()
    {
        throw new System.NotImplementedException();
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        if (this._streams.Count == 0)
        {
            return 0;
        }

        int totalBytesRead = 0;

        while (this._streams.Count > 0 && count > 0)
        {
            int bytesRead = this._streams.Peek().Read(buffer, offset, count);

            if (bytesRead == 0)
            {
                this._streams.Dequeue().Dispose();
                continue;
            }

            totalBytesRead += bytesRead;
            offset += bytesRead;
            count -= bytesRead;
        }

        this._position += totalBytesRead;
        return totalBytesRead;
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        throw new System.NotImplementedException();
    }

    public override void SetLength(long value)
    {
        throw new System.NotImplementedException();
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        throw new System.NotImplementedException();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            while (this._streams.Count > 0)
            {
                this._streams.Dequeue().Dispose();
            }

        }
        base.Dispose(disposing);
    }
}
