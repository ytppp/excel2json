#!/bin/bash

# 获取所有远程仓库名称
remotes=$(git remote)

# 遍历每个远程仓库并推送代码
for remote in $remotes; do
  git push $remote main
done