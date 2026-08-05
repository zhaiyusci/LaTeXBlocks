using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace LaTeXBlocks.RenderHost
{
    /// <summary>
    /// Version-one transport for one UTF-8 JSON object per length-prefixed frame.
    /// Requests are deliberately limited to 1 MiB. Responses can carry SVG bytes,
    /// so they have a separately bounded 8 MiB envelope; the render engine limits
    /// raw SVG to 5 MiB before base64 encoding, leaving room for JSON metadata.
    /// </summary>
    internal static class PipeProtocol
    {
        internal const int Version = 1;
        internal const int MaximumRequestFrameBytes = 1024 * 1024;
        internal const int MaximumResponseFrameBytes = 8 * 1024 * 1024;

        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false, true);
        private static readonly DataContractJsonSerializer RequestSerializer =
            new DataContractJsonSerializer(typeof(RenderHostRequest));
        private static readonly DataContractJsonSerializer ResponseSerializer =
            new DataContractJsonSerializer(typeof(RenderHostResponse));

        internal static RenderHostRequest ReadRequest(Stream stream)
        {
            var json = ReadFrame(stream, MaximumRequestFrameBytes, "request");
            if (json == null)
            {
                return null;
            }

            try
            {
                using (var input = new MemoryStream(Utf8.GetBytes(json), false))
                {
                    var request = (RenderHostRequest)RequestSerializer.ReadObject(input);
                    if (request == null)
                    {
                        throw new ProtocolException("invalid_request", "The request must be a JSON object.");
                    }

                    return request;
                }
            }
            catch (SerializationException exception)
            {
                throw new ProtocolException("invalid_request", "The request is not valid protocol JSON.", exception);
            }
        }

        internal static void WriteResponse(Stream stream, RenderHostResponse response)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if (response == null)
            {
                throw new ArgumentNullException(nameof(response));
            }

            using (var output = new MemoryStream())
            {
                ResponseSerializer.WriteObject(output, response);
                WriteFrame(stream, Utf8.GetString(output.ToArray()), MaximumResponseFrameBytes, "response");
            }
        }

        private static string ReadFrame(Stream stream, int maximumBytes, string kind)
        {
            var firstByte = stream.ReadByte();
            if (firstByte < 0)
            {
                return null;
            }

            var header = new byte[4];
            header[0] = (byte)firstByte;
            ReadExactly(stream, header, 1, 3);

            var length = (uint)(header[0]
                | (header[1] << 8)
                | (header[2] << 16)
                | (header[3] << 24));
            if (length > maximumBytes)
            {
                throw new ProtocolException(
                    "frame_too_large",
                    "The " + kind + " frame exceeds its protocol size limit.");
            }

            var payload = new byte[(int)length];
            ReadExactly(stream, payload, 0, payload.Length);
            try
            {
                return Utf8.GetString(payload);
            }
            catch (DecoderFallbackException exception)
            {
                throw new ProtocolException("invalid_encoding", "The request frame is not valid UTF-8.", exception);
            }
        }

        private static void WriteFrame(Stream stream, string json, int maximumBytes, string kind)
        {
            var payload = Utf8.GetBytes(json ?? string.Empty);
            if (payload.Length > maximumBytes)
            {
                throw new ProtocolException(
                    "response_too_large",
                    "The " + kind + " frame exceeds its protocol size limit.");
            }

            var length = payload.Length;
            var header = new[]
            {
                (byte)(length & 0xff),
                (byte)((length >> 8) & 0xff),
                (byte)((length >> 16) & 0xff),
                (byte)((length >> 24) & 0xff)
            };

            stream.Write(header, 0, header.Length);
            stream.Write(payload, 0, payload.Length);
            stream.Flush();
        }

        private static void ReadExactly(Stream stream, byte[] buffer, int offset, int count)
        {
            while (count > 0)
            {
                var read = stream.Read(buffer, offset, count);
                if (read <= 0)
                {
                    throw new EndOfStreamException("The pipe closed before its frame was complete.");
                }

                offset += read;
                count -= read;
            }
        }
    }

    [DataContract]
    internal sealed class RenderHostRequest
    {
        [DataMember(Name = "version", EmitDefaultValue = false)]
        internal int Version { get; set; }

        [DataMember(Name = "id", EmitDefaultValue = false)]
        internal string Id { get; set; }

        [DataMember(Name = "command", EmitDefaultValue = false)]
        internal string Command { get; set; }

        [DataMember(Name = "profile", EmitDefaultValue = false)]
        internal string Profile { get; set; }

        [DataMember(Name = "source", EmitDefaultValue = false)]
        internal string Source { get; set; }

        [DataMember(Name = "widthPt", EmitDefaultValue = false)]
        internal double? WidthPt { get; set; }

        [DataMember(Name = "autoWidth", EmitDefaultValue = false)]
        internal bool? AutoWidth { get; set; }

        [DataMember(Name = "fontSizePt", EmitDefaultValue = false)]
        internal double? FontSizePt { get; set; }
    }

    [DataContract]
    internal sealed class RenderHostResponse
    {
        [DataMember(Name = "version")]
        internal int Version { get; set; }

        [DataMember(Name = "id", EmitDefaultValue = false)]
        internal string Id { get; set; }

        [DataMember(Name = "ok")]
        internal bool Ok { get; set; }

        [DataMember(Name = "result", EmitDefaultValue = false)]
        internal RenderHostResult Result { get; set; }

        [DataMember(Name = "error", EmitDefaultValue = false)]
        internal RenderHostError Error { get; set; }
    }

    [DataContract]
    internal sealed class RenderHostResult
    {
        [DataMember(Name = "kind")]
        internal string Kind { get; set; }

        [DataMember(Name = "protocol", EmitDefaultValue = false)]
        internal string Protocol { get; set; }

        [DataMember(Name = "processId", EmitDefaultValue = false)]
        internal int ProcessId { get; set; }

        [DataMember(Name = "pipeName", EmitDefaultValue = false)]
        internal string PipeName { get; set; }

        [DataMember(Name = "status", EmitDefaultValue = false)]
        internal string Status { get; set; }

        [DataMember(Name = "profile", EmitDefaultValue = false)]
        internal string Profile { get; set; }

        [DataMember(Name = "profiles", EmitDefaultValue = false)]
        internal string[] Profiles { get; set; }

        [DataMember(Name = "svgBase64", EmitDefaultValue = false)]
        internal string SvgBase64 { get; set; }

        [DataMember(Name = "summaryJson", EmitDefaultValue = false)]
        internal string SummaryJson { get; set; }

        [DataMember(Name = "outcomeCode", EmitDefaultValue = false)]
        internal int OutcomeCode { get; set; }

        [DataMember(Name = "issueFlags", EmitDefaultValue = false)]
        internal int IssueFlags { get; set; }

        [DataMember(Name = "outcomeMessage", EmitDefaultValue = false)]
        internal string OutcomeMessage { get; set; }

        [DataMember(Name = "depthPt", EmitDefaultValue = false)]
        internal double DepthPt { get; set; }
    }

    [DataContract]
    internal sealed class RenderHostError
    {
        [DataMember(Name = "code")]
        internal string Code { get; set; }

        [DataMember(Name = "message")]
        internal string Message { get; set; }
    }

    internal sealed class ProtocolException : Exception
    {
        internal ProtocolException(string code, string message)
            : base(message)
        {
            Code = code;
        }

        internal ProtocolException(string code, string message, Exception innerException)
            : base(message, innerException)
        {
            Code = code;
        }

        internal string Code { get; }
    }
}
