# Signing PDFs

In this example we will show how to add a visible or invisible signature
to a generated PDF file.

## Concepts

- In order to sign a PDF file you will need a certificate issued by a
  valid Certificate Authority, or one issued by yourself. In this
  example we will use a self signed certificate. **This certificate
  will not validate by default when you open it in Acrobat, you need
  to add it to your trusted list.**

- The default algorithm for CmsSigner .NET class is SHA-1, which is
  known to have vulnerabilities and shouldn't be used anymore. So in
  this example we use SHA512 instead by changing the
  DigestAlgorithm.


- In order to sign a file, FlexCel **will write a requirement for
  Acrobat 8 or newer in the generated files. This is because only
  Acrobat 8 or newer support SHA512.** Older versions of acrobat
  will still display the pages but will not validate the signature.

- We provide a default signing implementation using standard .NET
  crypto classes**.** You can still create your own signature engine
  by using a third party cryptography library or by calling
  CryptoApi in windows via p/invoke. This is explained in the section
  [Signing PDF Files](https://doc.tmssoftware.com/flexcel/net/guides/pdf-exporting-guide.html#signing-pdf-files) in the 
  PDF exporting guide.

