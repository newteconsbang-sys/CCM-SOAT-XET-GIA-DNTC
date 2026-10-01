from pathlib import Path
from datetime import datetime,timedelta,timezone
from cryptography import x509
from cryptography.x509.oid import NameOID,ExtendedKeyUsageOID
from cryptography.hazmat.primitives import hashes,serialization
from cryptography.hazmat.primitives.asymmetric import rsa
import ipaddress
ROOT=Path(__file__).resolve().parents[1]
def create():
    folder=ROOT/'config';folder.mkdir(exist_ok=True)
    certpath=folder/'localhost.crt';keypath=folder/'localhost.key'
    if certpath.exists() and keypath.exists():
        cert=x509.load_pem_x509_certificate(certpath.read_bytes())
        if cert.not_valid_after_utc<datetime.now(timezone.utc)+timedelta(days=7):raise RuntimeError('Chứng chỉ hết hạn/sắp hết hạn; cần lưu bản cũ và tạo lại.')
        return
    key=rsa.generate_private_key(public_exponent=65537,key_size=2048)
    name=x509.Name([x509.NameAttribute(NameOID.COMMON_NAME,'CCM-AI localhost')])
    cert=(x509.CertificateBuilder().subject_name(name).issuer_name(name).public_key(key.public_key()).serial_number(x509.random_serial_number()).not_valid_before(datetime.now(timezone.utc)-timedelta(minutes=5)).not_valid_after(datetime.now(timezone.utc)+timedelta(days=365)).add_extension(x509.SubjectAlternativeName([x509.DNSName('localhost'),x509.IPAddress(ipaddress.ip_address('127.0.0.1'))]),critical=False).add_extension(x509.BasicConstraints(ca=False,path_length=None),critical=True).add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]),critical=False).sign(key,hashes.SHA256()))
    keypath.write_bytes(key.private_bytes(serialization.Encoding.PEM,serialization.PrivateFormat.PKCS8,serialization.NoEncryption()))
    certpath.write_bytes(cert.public_bytes(serialization.Encoding.PEM));(folder/'localhost.cer').write_bytes(cert.public_bytes(serialization.Encoding.DER))
if __name__=='__main__':create()
