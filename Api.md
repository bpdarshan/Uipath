## Code Snippet for Disabling the SSL Certification.

```vb.net

System.Net.ServicePointManager.ServerCertificateValidationCallback =
Function(se As Object,
cert As System.Security.Cryptography.X509Certificates.X509Certificate,
chain As System.Security.Cryptography.X509Certificates.X509Chain,
sslerror As System.Net.Security.SslPolicyErrors) True
