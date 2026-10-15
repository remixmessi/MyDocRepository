#!/bin/sh
# Invoked only by the installed app after explicit confirmation and successful save.
set -eu
parent_pid=$1
profile=$2
application=$3
case "$parent_pid" in ''|*[!0-9]*) exit 2;; esac
case "$profile" in "$HOME/Library/Application Support/MyDoc") ;; *) exit 2;; esac
case "$application" in /*/MyDoc.app) ;; *) exit 2;; esac
[ ! -L "$profile" ] && [ ! -L "$application" ] || exit 2
# A failed/blocked quit must never delete a running application's data.
tries=0
while kill -0 "$parent_pid" 2>/dev/null; do
  tries=$((tries + 1))
  [ "$tries" -lt 60 ] || exit 3
  sleep 1
done
if ! /bin/rm -rf -- "$profile" "$application"; then
  /usr/bin/osascript -e 'display alert "MyDoc 卸载未完成" message "部分文件无法删除，请检查应用与用户数据目录的权限。"'
  exit 1
fi
