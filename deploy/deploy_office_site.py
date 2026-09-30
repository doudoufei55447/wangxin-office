"""一键部署 office.wx12345.com：页面 + 安装包 + nginx 配置。

单次 SSH 会话内完成，减少连接次数（沙箱出网偶发抖动）：
  1. 备份服务器上的旧安装包 / README / nginx 站点配置
  2. 上传 site/index.html、download/*.msi、download/*.dmg
  3. certutil 校验远端 SHA-256 与本机一致
  4. 替换 nginx 站点配置 → nginx -t → 通过才 reload（失败自动回滚）
  5. 删除 Web 根下暴露的旧 v1.0.4 安装包与内部 README.txt

用法: python deploy_office_site.py [--dry-run]
"""
import hashlib
import os
import posixpath
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import remote  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

SITE_DIR = os.path.join(ROOT, 'site')
UPLOAD_DIR = os.path.join(HERE, '_upload')
BACKUP_DIR = os.path.join(HERE, '_remote', 'backup')

REMOTE_SITE = 'C:/nginx/nginx-1.26.2/html/office.wx12345.com'
REMOTE_DOWNLOAD = REMOTE_SITE + '/download'
REMOTE_CONF_DIR = 'C:/nginx/nginx-1.26.2/conf/sites-enabled'
NGINX_HOME = r'C:\nginx\nginx-1.26.2'
NGINX_EXE = NGINX_HOME + r'\nginx.exe'

LOCAL_NGINX_CONF = os.path.join(HERE, 'nginx', 'office.wx12345.com.conf')
REMOTE_NGINX_CONF = REMOTE_CONF_DIR + '/office.wx12345.com.conf'

PAGES = ['index.html', 'logo.png', 'favicon.svg']
PACKAGES = ['wangxin-office-1.0.8-x64.msi', 'wangxin-office-1.0.8-arm64.dmg']
STALE = [REMOTE_DOWNLOAD + '/wangxin-office-setup-win.exe',
         REMOTE_DOWNLOAD + '/README.txt']


def sha256_local(path):
    h = hashlib.sha256()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def sha256_remote(ssh, rpath):
    out, _err = remote.run(ssh, 'certutil -hashfile "%s" SHA256' % rpath, timeout=120)
    for line in out.splitlines():
        token = line.strip().replace(' ', '')
        if len(token) == 64 and all(c in '0123456789abcdefABCDEF' for c in token):
            return token.lower()
    return None


def main():
    dry = '--dry-run' in sys.argv
    os.makedirs(BACKUP_DIR, exist_ok=True)

    ssh = remote.connect()
    try:
        sftp = ssh.open_sftp()

        # ---- 1. 备份 ----
        print('== 1. 备份服务器现有文件 ==')
        for rpath in [REMOTE_SITE + '/index.html',
                      REMOTE_NGINX_CONF] + STALE:
            name = posixpath.basename(rpath)
            dst = os.path.join(BACKUP_DIR, name)
            try:
                sftp.get(rpath, dst)
                print('   backed up %-38s -> %s' % (name, os.path.relpath(dst, HERE)))
            except IOError:
                print('   (skip, not found) %s' % name)

        # ---- 2. 上传 ----
        print('== 2. 上传页面与安装包 ==')
        for name in PAGES:
            lp = os.path.join(SITE_DIR, name)
            sftp.put(lp, REMOTE_SITE + '/' + name)
            print('   PUT %-34s %d bytes' % (name, os.path.getsize(lp)))
        for name in PACKAGES:
            lp = os.path.join(UPLOAD_DIR, name)
            sftp.put(lp, REMOTE_DOWNLOAD + '/' + name)
            print('   PUT %-34s %d bytes' % (name, os.path.getsize(lp)))

        # ---- 3. 校验哈希 ----
        print('== 3. 校验远端 SHA-256 ==')
        ok = True
        for name in PACKAGES:
            local_h = sha256_local(os.path.join(UPLOAD_DIR, name))
            remote_h = sha256_remote(ssh, REMOTE_DOWNLOAD + '/' + name)
            match = (local_h == remote_h)
            ok = ok and match
            print('   %-34s %s' % (name, 'MATCH' if match else 'MISMATCH 本地=%s 远端=%s' % (local_h, remote_h)))
        if not ok:
            sys.exit('远端哈希不一致，已中止（nginx 配置未改动）')

        # ---- 4. nginx 配置 ----
        print('== 4. 更新 nginx 站点配置 ==')
        if dry:
            print('   (dry-run) 跳过')
        else:
            sftp.put(LOCAL_NGINX_CONF, REMOTE_NGINX_CONF)
            print('   PUT office.wx12345.com.conf')
            out, err = remote.run(
                ssh, 'cd /d %s && "%s" -p %s -t' % (NGINX_HOME, NGINX_EXE, NGINX_HOME), timeout=120)
            combined = (out + '\n' + err)
            print('   nginx -t:')
            for line in combined.strip().splitlines():
                print('     ' + line)
            if 'successful' not in combined:
                print('   !! 配置测试失败，回滚旧配置')
                sftp.put(os.path.join(BACKUP_DIR, 'office.wx12345.com.conf'), REMOTE_NGINX_CONF)
                sys.exit('nginx -t 失败，已回滚')
            out, err = remote.run(ssh, 'cd /d %s && "%s" -p %s -s reload' % (NGINX_HOME, NGINX_EXE, NGINX_HOME), timeout=120)
            print('   reload:', (out + err).strip() or '(no output = ok)')

        # ---- 5. 清理旧文件 ----
        print('== 5. 清理 Web 根下的旧文件 ==')
        for rpath in STALE:
            if dry:
                print('   (dry-run) would delete %s' % rpath)
                continue
            out, err = remote.run(ssh, 'del /f /q "%s"' % rpath.replace('/', '\\'), timeout=60)
            print('   del %-46s %s' % (posixpath.basename(rpath), (out + err).strip() or 'ok'))

        sftp.close()
        print('\nDONE')
    finally:
        ssh.close()


if __name__ == '__main__':
    main()
