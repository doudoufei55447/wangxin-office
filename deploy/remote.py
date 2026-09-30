"""网信办公 office.wx12345.com 部署通道（paramiko over SSH）。

用法:
  python remote.py probe                       # 连通性 + 服务器清单
  python remote.py sh --cmd "dir C:\\inetpub"   # 在服务器上跑一条 cmd 命令
  python remote.py put <local> <remote>        # 上传文件
  python remote.py get <remote> <local>        # 下载文件

口令来源（按优先级）: 环境变量 WM_SRV_PASS / OFFICE_SRV_PASS
                     → D:\\Web应用监控\\.env → deploy/.env
"""
import os
import sys
import time
import socket
import argparse

import paramiko

HOST = os.environ.get('OFFICE_SRV_HOST', '62.234.156.247')
PORT = 22
USER = os.environ.get('OFFICE_SRV_USER', 'Administrator')

ENV_FILES = [
    r'D:\Web应用监控\.env',
    os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env'),
]


def read_pass():
    pw = os.environ.get('WM_SRV_PASS') or os.environ.get('OFFICE_SRV_PASS')
    if pw:
        return pw
    for path in ENV_FILES:
        try:
            with open(path, encoding='utf-8') as fh:
                for line in fh:
                    line = line.strip()
                    for key in ('WM_SRV_PASS=', 'OFFICE_SRV_PASS='):
                        if line.startswith(key):
                            return line.split('=', 1)[1].strip().strip('"').strip("'")
        except OSError:
            continue
    sys.exit('未找到服务器口令：请设置环境变量 WM_SRV_PASS 或在 deploy/.env 中提供')


def decode(raw):
    """服务器中文 locale 输出为 GBK，先 UTF-8 再 cp936。"""
    for enc in ('utf-8', 'cp936'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', errors='replace')


def connect(retries=3):
    s = socket.socket()
    s.settimeout(5)
    try:
        s.connect((HOST, PORT))
    except Exception:
        sys.exit(f'TCP {HOST}:{PORT} 不可达')
    finally:
        s.close()

    last = None
    for attempt in range(1, retries + 1):
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        try:
            ssh.connect(HOST, port=PORT, username=USER, password=read_pass(),
                        timeout=30, banner_timeout=30, auth_timeout=30,
                        look_for_keys=False, allow_agent=False)
            return ssh
        except Exception as exc:          # 瞬时 banner 读取失败 → 重试
            last = exc
            try:
                ssh.close()
            except Exception:
                pass
            if attempt < retries:
                time.sleep(3 * attempt)
    sys.exit(f'SSH 连接失败（{retries} 次重试后）: {last!r}')


def run(ssh, cmd, timeout=300):
    _in, out, err = ssh.exec_command(cmd, timeout=timeout)
    o = decode(out.read())
    e = decode(err.read())
    return o, e


def cmd_probe(ssh):
    checks = [
        ('hostname', 'hostname'),
        ('os', 'wmic os get Caption,Version /value'),
        ('iis', 'C:\\Windows\\System32\\inetsrv\\appcmd.exe list sites'),
        ('inetpub', 'dir /b C:\\inetpub'),
        ('disk', 'wmic logicaldisk get DeviceID,FreeSpace,Size /value'),
        ('outbound-https', 'curl.exe -s -o NUL -w "%{http_code}" https://api.github.com'),
    ]
    for label, c in checks:
        o, e = run(ssh, c, timeout=90)
        print(f'--- {label} ---')
        print((o or '').strip() or '(empty)')
        if e.strip():
            print('  [stderr]', e.strip()[:400])
        print()


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='mode', required=True)

    sub.add_parser('probe')

    p_sh = sub.add_parser('sh')
    p_sh.add_argument('--cmd', required=True)
    p_sh.add_argument('--timeout', type=int, default=300)

    p_put = sub.add_parser('put')
    p_put.add_argument('local')
    p_put.add_argument('remote')

    p_get = sub.add_parser('get')
    p_get.add_argument('remote')
    p_get.add_argument('local')

    p_ls = sub.add_parser('ls')
    p_ls.add_argument('remote')

    p_pull = sub.add_parser('pull')
    p_pull.add_argument('remote_dir')
    p_pull.add_argument('local_dir')
    p_pull.add_argument('--files', nargs='*', help='缺省取回该目录全部普通文件')

    args = ap.parse_args()
    ssh = connect()
    try:
        if args.mode == 'probe':
            cmd_probe(ssh)
        elif args.mode == 'sh':
            o, e = run(ssh, args.cmd, timeout=args.timeout)
            print(o)
            if e.strip():
                print('=== STDERR ===')
                print(e)
        else:
            sftp = ssh.open_sftp()
            try:
                if args.mode == 'put':
                    sftp.put(args.local, args.remote)
                    print('PUT', args.local, '->', args.remote)
                elif args.mode == 'get':
                    sftp.get(args.remote, args.local)
                    print('GET', args.remote, '->', args.local)
                elif args.mode == 'pull':
                    os.makedirs(args.local_dir, exist_ok=True)
                    names = args.files
                    if not names:
                        names = [a.filename for a in sftp.listdir_attr(args.remote_dir)
                                 if not a.st_mode & 0o040000]
                    for name in names:
                        src = args.remote_dir.rstrip('/') + '/' + name
                        dst = os.path.join(args.local_dir, name)
                        sftp.get(src, dst)
                        print(f'GET {name} ({os.path.getsize(dst)} bytes)')
                else:
                    for a in sorted(sftp.listdir_attr(args.remote), key=lambda x: x.filename):
                        kind = 'DIR ' if a.st_mode & 0o040000 else '    '
                        print(f'{kind}{a.st_size:>12}  {a.filename}')
            finally:
                sftp.close()
    finally:
        ssh.close()


if __name__ == '__main__':
    main()
