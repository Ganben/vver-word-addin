# Word Add-In: Visual Verified for signed/encrypted document

## Motivation

a word add-in, use common-API and Word-API to:

1 calculate in page paragraph hash merk-trees
2 represent in QR code and insert base64 img/finger prints to each page and master node page
3 can verify in page - data/finger prints
paras(utf-8 coding) in page --> page hash(img/finger prints hex)
page hashes --> document hash(img/finger prints hex)

### How dose it work


ECC sig:
> 64 bytes
> For example, for 256-bit elliptic curves (like secp256k1 ) the ECDSA signature is 512 bits (64 bytes) and for 521-bit curves (like secp521r1 ) the signature is 1042 bits.```

QR code size:
> 406 bytes
> A 101x101 QR code, with high level error correction, can hold 3248 bits, or 406 bytes. Probably not enough for any meaningful SVG/XML data. A 177x177 grid, depending on desired level of error correction, can store between 1273 and 2953 bytes.



### visual-verifier dev notes

[Source: yeoman generator](https://docs.microsoft.com/en-us/office/dev/add-ins/quickstarts/word-quickstart?tabs=yeomangenerator)

`cd visual-verifier`

`npm start`
PS: need a PS command turn on web settings ? add later



#### deps

npm:

qrcode
merkletreejs
merkle-tools

