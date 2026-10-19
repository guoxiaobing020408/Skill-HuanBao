#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用 GitHub REST API 把本地目录整体同步到仓库（不需要 git 协议，只走 api.github.com）。

适用于 github.com 网页/443 被网络策略挡住、但 api.github.com 可访问的环境。

用法：
    python push_api.py --repo 用户名/仓库名 --dir . --message "update"

Token 来源优先级：
    1) --token 参数
    2) 环境变量 GITHUB_TOKEN
    3) 文件 %USERPROFILE%\\.workbuddy\\github_token.txt

依赖：仅标准库。
"""

import argparse
import base64
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request

API = 'https://api.github.com'
SKIP_DIRS = {'.git', '__pycache__', 'out', '.idea', '.vscode'}
SKIP_FILES = {'.token', '.DS_Store', 'Thumbs.db', 'desktop.ini'}
SKIP_SUFFIX = ('.pyc', '.pyo')


def token_path():
    return os.path.join(os.path.expanduser('~'), '.workbuddy', 'github_token.txt')


def load_token(arg):
    if arg:
        return arg.strip()
    env = (os.environ.get('GITHUB_TOKEN') or '').strip()
    if env:
        return env
    p = token_path()
    if os.path.exists(p):
        with open(p, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return ''


def call(method, path, token, payload=None, ok=(200, 201)):
    url = API + path
    data = None
    headers = {
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
        'User-Agent': 'guo6-skill-pusher',
        'Authorization': 'Bearer ' + token,
    }
    if payload is not None:
        data = json.dumps(payload).encode('utf-8')
        headers['Content-Type'] = 'application/json'
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            body = r.read().decode('utf-8')
            return r.status, (json.loads(body) if body.strip() else {})
    except urllib.error.HTTPError as e:
        body = e.read().decode('utf-8', 'replace')
        try:
            parsed = json.loads(body)
        except Exception:
            parsed = {'message': body[:400]}
        return e.code, parsed


def collect(root):
    """返回 [(相对路径, 绝对路径)]，路径统一用 / 分隔，跳过无关文件。"""
    out = []
    for cur, dirs, files in os.walk(root):
        dirs[:] = sorted(d for d in dirs if d not in SKIP_DIRS)
        for name in sorted(files):
            if name in SKIP_FILES or name.lower().endswith(SKIP_SUFFIX):
                continue
            full = os.path.join(cur, name)
            rel = os.path.relpath(full, root).replace(os.sep, '/')
            out.append((rel, full))
    return out


def is_text(path):
    try:
        with open(path, 'rb') as f:
            chunk = f.read(8000)
    except OSError:
        return False
    if b'\x00' in chunk:
        return False
    try:
        chunk.decode('utf-8')
        return True
    except UnicodeDecodeError:
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--repo', required=True, help='用户名/仓库名')
    ap.add_argument('--dir', default='.', help='要上传的本地目录')
    ap.add_argument('--branch', default='main')
    ap.add_argument('--message', default='update skill')
    ap.add_argument('--token', default='')
    ap.add_argument('--dry-run', action='store_true')
    args = ap.parse_args()

    token = load_token(args.token)
    if not token:
        print('ERROR: 没有找到 GitHub Token。')
        print('  把 token 写进 %s ，或用 --token 传入。' % token_path())
        return 2

    root = os.path.abspath(args.dir)

    status, me = call('GET', '/user', token)
    if status != 200:
        print('ERROR: Token 无效（HTTP %s）：%s' % (status, me.get('message')))
        print('  常见原因：token 过期、拼错、缺少 repo 权限。')
        return 2
    login = me.get('login', '?')

    status, repo = call('GET', '/repos/' + args.repo, token)
    if status != 200:
        print('ERROR: 访问仓库 %s 失败（HTTP %s）：%s' % (args.repo, status, repo.get('message')))
        return 2
    print('账号 %s  ->  仓库 %s（%s）' % (login, repo['full_name'],
                                        'public' if not repo['private'] else 'private'))

    # 现有分支的父提交与基准树
    parent = None
    status, ref = call('GET', '/repos/%s/git/ref/heads/%s' % (args.repo, args.branch), token)
    if status == 200:
        parent = ref['object']['sha']
        print('现有分支 %s 指向 %s' % (args.branch, parent[:7]))
    elif status in (404, 409):
        print('分支 %s 还不存在，将新建（空仓库首次推送）' % args.branch)
    else:
        print('WARN: 读分支失败（HTTP %s）：%s' % (status, ref.get('message')))

    files = collect(root)
    if not files:
        print('ERROR: %s 里没有可上传的文件' % root)
        return 2

    # 空仓库无法直接建 blob（HTTP 409 Git Repository is empty），
    # 先用 Contents API 塞一个占位文件把仓库「点燃」，最后再把引用指到真正的提交上。
    bootstrapped = False
    if parent is None and repo.get('size', 0) == 0:
        if args.dry_run:
            print('仓库为空，正式上传时会先做一次占位提交。')
        else:
            seed = files[0][0]
            with open(files[0][1], 'rb') as f:
                seed_b64 = base64.b64encode(f.read()).decode('ascii')
            st, _ = call('PUT', '/repos/%s/contents/%s' % (args.repo, urllib.parse.quote(seed)),
                         token, {'message': 'init', 'content': seed_b64, 'branch': args.branch})
            if st not in (200, 201):
                print('ERROR: 初始化空仓库失败（HTTP %s）：%s' % (st, _.get('message')))
                return 1
            bootstrapped = True
            print('空仓库已初始化（占位文件 %s，稍后会被真实内容覆盖）' % seed)

    print('待上传 %d 个文件：' % len(files))
    for rel, _ in files:
        print('  ' + rel)
    if args.dry_run:
        print('\n--dry-run：只列出，不写入。')
        return 0

    # 1) 建 blob
    tree = []
    for rel, full in files:
        with open(full, 'rb') as f:
            raw = f.read()
        if is_text(full):
            payload = {'content': raw.decode('utf-8'), 'encoding': 'utf-8'}
        else:
            payload = {'content': base64.b64encode(raw).decode('ascii'), 'encoding': 'base64'}
        st, blob = call('POST', '/repos/%s/git/blobs' % args.repo, token, payload)
        if st != 201:
            print('ERROR: 上传 %s 失败（HTTP %s）：%s' % (rel, st, blob.get('message')))
            return 1
        tree.append({'path': rel, 'mode': '100644', 'type': 'blob', 'sha': blob['sha']})

    # 2) 建树（不给 base_tree = 整棵树覆盖，即完全镜像本地目录）
    st, new_tree = call('POST', '/repos/%s/git/trees' % args.repo, token, {'tree': tree})
    if st != 201:
        print('ERROR: 建树失败（HTTP %s）：%s' % (st, new_tree.get('message')))
        return 1

    # 3) 建提交
    commit_payload = {'message': args.message, 'tree': new_tree['sha']}
    if parent and not bootstrapped:
        commit_payload['parents'] = [parent]
    st, commit = call('POST', '/repos/%s/git/commits' % args.repo, token, commit_payload)
    if st != 201:
        print('ERROR: 建提交失败（HTTP %s）：%s' % (st, commit.get('message')))
        return 1
    print('提交：%s  %s' % (commit['sha'][:7], args.message))

    # 4) 更新或新建分支引用
    if parent or bootstrapped:
        st, res = call('PATCH', '/repos/%s/git/refs/heads/%s' % (args.repo, args.branch),
                       token, {'sha': commit['sha'], 'force': True})
    else:
        st, res = call('POST', '/repos/%s/git/refs' % args.repo, token,
                       {'ref': 'refs/heads/' + args.branch, 'sha': commit['sha']})
    if st not in (200, 201):
        print('ERROR: 更新分支失败（HTTP %s）：%s' % (st, res.get('message')))
        return 1

    print('\n完成：https://github.com/%s/tree/%s' % (args.repo, args.branch))
    return 0


if __name__ == '__main__':
    sys.exit(main())
