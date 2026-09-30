"""office.wx12345.com 公网验收：裸 socket 直连，绕过代理。

检查项：
  - HTTPS 首页 200 且含 v1.0.8 与两端下载链接
  - 两个安装包 200 + Content-Length 正确 + Content-Disposition: attachment
  - HTTP 301 跳 HTTPS
  - 证书主体 / 剩余天数
  - 旧文件（v1.0.4 exe / README.txt）已 404
"""
import re
import socket
import ssl
import sys

DOMAIN = 'office.wx12345.com'
IP = socket.gethostbyname(DOMAIN)
EXPECT = {
    '/download/wangxin-office-1.0.8-x64.msi': 2646016,
    '/download/wangxin-office-1.0.8-arm64.dmg': 3460587,
}
GONE = ['/download/wangxin-office-setup-win.exe', '/download/README.txt']

fails = []


def request(port, target, use_tls):
    raw = socket.create_connection((IP, port), timeout=20)
    if use_tls:
        ctx = ssl.create_default_context()
        raw = ctx.wrap_socket(raw, server_hostname=DOMAIN)
    req = ('GET %s HTTP/1.1\r\nHost: %s\r\nConnection: close\r\n'
           'User-Agent: verify\r\n\r\n') % (target, DOMAIN)
    raw.sendall(req.encode())
    buf = b''
    while True:
        chunk = raw.recv(65536)
        if not chunk:
            break
        buf += chunk
    raw.close()
    head, _, body = buf.partition(b'\r\n\r\n')
    return head.decode('latin-1'), body


def parse(head):
    lines = head.split('\r\n')
    status = int(lines[0].split()[1])
    headers = {}
    for line in lines[1:]:
        if ':' in line:
            k, v = line.split(':', 1)
            headers[k.strip().lower()] = v.strip()
    return status, headers


# --- 证书 ---
ctx = ssl.create_default_context()
with socket.create_connection((IP, 443), timeout=20) as s:
    with ctx.wrap_socket(s, server_hostname=DOMAIN) as ss:
        cert = ss.getpeercert()
print('证书主体 :', dict(x[0] for x in cert['subject']).get('commonName'))
print('证书 SAN  :', [v for k, v in cert['subjectAltName'] if k == 'DNS'])
print('证书到期 :', cert['notAfter'])
print()

# --- 首页 ---
head, body = request(443, '/', True)
status, headers = parse(head)
html = body.decode('utf-8', 'replace')
print('GET / ->', status, headers.get('content-type'))
if status != 200:
    fails.append('首页非 200: %s' % status)
for token in ['v1.0.8', 'wangxin-office-1.0.8-x64.msi', 'wangxin-office-1.0.8-arm64.dmg',
              '已签名', 'beian.miit.gov.cn']:
    if token not in html:
        fails.append('首页缺少关键内容: %s' % token)
print('  首页含 v1.0.8 / 两个下载链接 / 备案链接 :', 'OK' if not fails else '见失败项')
print()

# --- 安装包 ---
for path, size in EXPECT.items():
    head, body = request(443, path, True)
    status, headers = parse(head)
    length = headers.get('content-length')
    disp = headers.get('content-disposition', '')
    print('GET %s' % path)
    print('   ->', status, '| Content-Length:', length, '| Content-Disposition:', disp or '(无)')
    if status != 200:
        fails.append('%s 非 200: %s' % (path, status))
    if length and int(length) != size:
        fails.append('%s 大小不符 期望 %d 实际 %s' % (path, size, length))
    if 'attachment' not in disp.lower():
        fails.append('%s 缺少 Content-Disposition: attachment' % path)
print()

# --- 旧文件应 404 ---
for path in GONE:
    head, _ = request(443, path, True)
    status, _h = parse(head)
    print('旧文件 %s -> %s' % (path, status))
    if status != 404:
        fails.append('%s 旧文件仍可访问: %s' % (path, status))
print()

# --- HTTP 跳转 ---
head, _ = request(80, '/download/wangxin-office-1.0.8-x64.msi', False)
status, headers = parse(head)
loc = headers.get('location', '')
print('HTTP 301 检查 ->', status, '| Location:', loc)
if status != 301 or not loc.startswith('https://'):
    fails.append('HTTP 未正确跳转 HTTPS: %s %s' % (status, loc))
if 'wangxin-office-1.0.8-x64.msi' not in loc:
    fails.append('301 未保留路径: %s' % loc)
print()

if fails:
    print('FAILED:')
    for f in fails:
        print('  -', f)
    sys.exit(1)
print('ALL CHECKS PASSED')
