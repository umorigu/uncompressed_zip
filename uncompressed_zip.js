// This source code is in the public domain.

var window = this;
(function(){
    var adTypeBinary = 1;
    var adReadAll = -1;
    var adSaveCreateOverWrite = 2;

    var fso = new ActiveXObject("Scripting.FileSystemObject");

    function convertUTF16StringToByteArray(u16str) {
        var u8str = unescape(encodeURIComponent(u16str));
        var ba = new ByteArray();
        for (var i = 0; i < u8str.length; i++) {
            ba.append(u8str.charCodeAt(i), 1);
        }
        return ba;
    }
    function convertBinaryBlockToByteArray(binBlock, size) {
        var ms = new ActiveXObject("System.IO.MemoryStream");
        ms.Write(binBlock, 0, size);
        ms.Position = 0;
        var ba = new ByteArray();
        for (var i = 0; i < size; i++) {
            var b = ms.ReadByte();
            ba.append(b, 1);
        }
        return ba;
    }
    function isAllAscii(byteArray) {
        var b = true;
        for (var i = 0; i < byteArray.length; i++) {
            var c = byteArray[i];
            if (c < 0x00 || 0x7f < c) {
                b = false;
                break;
            }
        }
        return b;
    }

    function Zip() {
        this.members = [];
    }
    Zip.prototype = {
        addMember: function (member) {
            this.members.push(member);
            return member;
        },
        addFile: function (rpath) {
            var name = rpath.replace(/\\/g, "/");
            var stream = new ActiveXObject('ADODB.Stream');
            stream.Type = adTypeBinary;
            stream.Open();
            stream.LoadFromFile(rpath);
            var allBytes = stream.Read(adReadAll);
            stream.Close();

            var f = fso.GetFile(rpath);
            var size = f.Size;
            
            var dt = new Date(f.DateLastModified);
            return this.addMember(new StringMember(allBytes, size, name, dt));
        },
        getBlockArray: function () {
            var blockArray = [];
            var members = this.members;
            var push = Array.prototype.push;
            var offsets = [];

            var curpos = 0;
            for (var i = 0; i < members.length; i++) {
                offsets.push(curpos);
                var ba = members[i].getLocalFileHeader();
                blockArray.push({"byte_array" : ba });
                curpos += ba.length;

                // file body
                blockArray.push( {
                        "binary_block" : members[i].getData(),
                        "binary_size" : members[i].dataLength }
                    );
                curpos += members[i].dataLength;
            }

            var bin = new ByteArray;
            var centralDirectoryOffset = curpos;

            for (var i = 0; i < members.length; i++) {
                push.apply(bin, members[i].getCentralDirectoryFileHeader(offsets[i]));
            }

            var endOfCentralDirectoryOffset = bin.length + centralDirectoryOffset;

//          end of central dir signature    4 bytes  (0x06054b50)
            bin.append(0x06054b50, 4);
//          number of this disk             2 bytes
            bin.append(0, 2);
//          number of the disk with the
//          start of the central directory  2 bytes
            bin.append(0, 2);
//          total number of entries in the
//          central directory on this disk  2 bytes
            bin.append(members.length, 2);
//          total number of entries in
//          the central directory           2 bytes
            bin.append(members.length, 2);
//          size of the central directory   4 bytes
            bin.append(endOfCentralDirectoryOffset - centralDirectoryOffset, 4);
//          offset of start of central
//          directory with respect to
//          the starting disk number        4 bytes
            bin.append(centralDirectoryOffset, 4);
//          .ZIP file comment length        2 bytes
            bin.append(0, 2);
//          .ZIP file comment       (variable size)
//          Array.prototype.push.apply(bin, []);

            blockArray.push({"byte_array" : bin });

            return blockArray;
        },
        saveFile: function (filename) {
            var blockArray = this.getBlockArray();
            var stream = new ActiveXObject('ADODB.Stream');
            stream.Type = adTypeBinary;
            stream.Open();

            for (var i = 0; i < blockArray.length; i++) {
                var block = blockArray[i];
                var byteArray = block["byte_array"];
                if (byteArray) {
                    var ms = new ActiveXObject("System.IO.MemoryStream");
                    for (var j = 0; j < byteArray.length; j++) {
                        ms.WriteByte(byteArray[j]);
                    }
                    stream.Write(ms.ToArray());
                } else {
                    var bin = block["binary_block"];
                    stream.Write(bin);
                }
            }
            stream.SaveToFile(filename, adSaveCreateOverWrite);
            stream.Close();
        },
        constructor: Zip
    };

    var crc32table = function() {
        var poly = 0xEDB88320, u, table = [];
        for (var i = 0; i < 256; i ++) {
            u = i;
            for (var j = 0; j < 8; j++) {
                if (u & 1)
                    u = (u >>> 1) ^ poly;
                else
                    u = u >>> 1;
            }
            table[i] = u;
        }
        return table;
    } ();

    var getCrc32 = function(bin) {
        var result = 0xFFFFFFFF;
        for (var i = 0; i < bin.length; i ++)
            result = (result >>> 8) ^ crc32table[bin[i] ^ (result & 0xFF)];
        return ~result;
    };

    function ByteArray() {
        var self = [];
        var proto = ByteArray.prototype;
        for (var name in proto)
            self[name] = proto[name];
        return self;
    }
    ByteArray.prototype = {
        append: function(value, bytes) {
            for (var i = 0; i < bytes; i ++)
                this.push(value >> (i * 8) & 0xFF);
        },
        constructor: ByteArray
    };

    function Member() { }
    Member.prototype = {
        initDateTime: function(dt) {
            this.date = ((dt.getFullYear() - 1980) << 9) |
                        ((dt.getMonth() + 1) << 5) |
                        (dt.getDate());
            this.time = (dt.getHours() << 5) |
                        (dt.getMinutes() << 5) |
                        (dt.getSeconds() >> 1);
        },
        getLocalFileHeader: function() {
            var bin = new ByteArray();
//          local file header signature     4 bytes  (0x04034b50)
            bin.append(0x04034b50, 4);
//          version needed to extract       2 bytes
            bin.append(10, 2);
//          general purpose bit flag        2 bytes
            var flag = 0;
            if (! isAllAscii(this.name)) {
                flag = 1 << 11;
            }
            bin.append(flag, 2);
//          compression method              2 bytes
            bin.append(0, 2);
//          last mod file time              2 bytes
            bin.append(this.time, 2);
//          last mod file date              2 bytes
            bin.append(this.date, 2);
//          crc-32                          4 bytes
            bin.append(this.crc32, 4);
//          compressed size                 4 bytes
            bin.append(this.dataLength, 4);
//          uncompressed size               4 bytes
            bin.append(this.dataLength, 4);
//          file name length                2 bytes
            bin.append(this.name.length, 2);
//          extra field length              2 bytes
            bin.append(this.extra.localFile.length, 2);
//          file name (variable size)
            Array.prototype.push.apply(bin, this.name);
//          extra field (variable size)
            Array.prototype.push.apply(bin, this.extra.localFile);
            return bin;
        },
        getData: function() {
            return this.data;
        },
        getCentralDirectoryFileHeader: function(offset) {
            var bin = new ByteArray();
//          central file header signature   4 bytes  (0x02014b50)
            bin.append(0x02014b50, 4);
//          version made by                 2 bytes
            bin.append(0x0317, 2);
//          version needed to extract       2 bytes
            bin.append(10, 2);
//          general purpose bit flag        2 bytes
            var flag = 0;
            if (! isAllAscii(this.name)) {
                flag = 1 << 11;
            }
            bin.append(flag, 2);
//          compression method              2 bytes
            bin.append(0, 2);
//          last mod file time              2 bytes
            bin.append(this.time, 2);
//          last mod file date              2 bytes
            bin.append(this.date, 2);
//          crc-32                          4 bytes
            bin.append(this.crc32, 4);
//          compressed size                 4 bytes
            bin.append(this.dataLength, 4);
//          uncompressed size               4 bytes
            bin.append(this.dataLength, 4);
//          file name length                2 bytes
            bin.append(this.name.length, 2);
//          extra field length              2 bytes
            bin.append(this.extra.centralDirectory.length, 2);
//          file comment length             2 bytes
            bin.append(0, 2);
//          disk number start               2 bytes
            bin.append(0, 2);
//          internal file attributes        2 bytes
            bin.append(0, 2);
//          external file attributes        4 bytes
            bin.append(this.externalFileAttributes, 4);
//          relative offset of local header 4 bytes
            bin.append(offset, 4);
//          file name (variable size)
            Array.prototype.push.apply(bin, this.name);
//          extra field (variable size)
            Array.prototype.push.apply(bin, this.extra.centralDirectory);
//          file comment (variable size)
//          Array.prototype.push.apply(bin, []);
            return bin;
        },
        constructor: Member
    };

    function Extra() {
        this.localFile = new ByteArray;
        this.centralDirectory = new ByteArray;
    }
    Extra.prototype = {
        append: function(field) {
            Array.prototype.push.apply(
                this.localFile,
                field.localFile
            );
            Array.prototype.push.apply(
                this.centralDirectory,
                field.centralDirectory
            );
        },
        contructor: ExtraField
    };

    function ExtraField(magic) {
        this.magic = magic;
    }

    function StringMember(binaryBlock, dataLength, name, date) {
        this.name = convertUTF16StringToByteArray(name);
        this.data = binaryBlock;
        this.dataLength = dataLength;
        this.crc32 = getCrc32(convertBinaryBlockToByteArray(binaryBlock, dataLength));
        this.externalFileAttributes = 0100644 << 16;
        this.initDateTime(new Date);
        this.date = ((date.getFullYear() - 1980) << 9) |
            ((date.getMonth() + 1) << 5) |
            (date.getDate());
        this.time = (date.getHours() << 5) |
            (date.getMinutes() << 5) |
            (date.getSeconds() >> 1);
        this.extra = new Extra;
    }
    StringMember.prototype = new Member();
    StringMember.constructor = StringMember;

    window.Zip = Zip;
})();

(function main() {
    var params = WScript.Arguments;
    if (params.length < 2) {
        WScript.StdErr.WriteLine("Error:");
        WScript.StdErr.WriteLine("Incorrect command line");
        WScript.StdErr.WriteLine("Usage: zip.js <archive_name> [<file_names>...]");
        return;
    }

    var fso = new ActiveXObject("Scripting.FileSystemObject");

    var archiveFilePath = params(0);

    WScript.StdOut.WriteLine("Creating archive " + archiveFilePath);
    WScript.StdOut.WriteLine();

    var folderPath = params(1);
    var folder = fso.GetFolder(folderPath);
    var rootFolderPath = folder.Path.substr(0, folder.Path.length - folderPath.length);
    var rootFolder = fso.GetFolder(rootFolderPath);

    var zip = new Zip;
    function Listing(folder) {
        for (var fc = new Enumerator(folder.Files); !fc.atEnd(); fc.moveNext()) {
            var f = fc.item();
            var rPath = f.Path.substr(rootFolder.Path.length + 1);
            zip.addFile(rPath);
            WScript.StdOut.WriteLine("Compressing  " + rPath);
        }
        for (var fc = new Enumerator(folder.SubFolders); !fc.atEnd(); fc.moveNext()) {
            Listing(fc.item());
        }
    }
    Listing(folder);
    WScript.StdOut.WriteLine();

    zip.saveFile(archiveFilePath);
    WScript.StdOut.WriteLine("Everything Ok");
})();
