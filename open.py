import base64
code = b'TXkgZ21haWwgYWNjb3VudCBpcyAtPiBtb2lzZXMuZ3JhdGVyb2wuMTNAZ21haWwuY29t'
byte_ary = base64.b64decode(code)
print(byte_ary.decode("ascii"))





