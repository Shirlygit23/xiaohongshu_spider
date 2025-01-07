# 每日开发检查表

## 开始开发前
- [ ] 拉取最新代码：`git pull`
- [ ] 检查当前分支：`git status`
- [ ] 确认开发任务和目标

## 开发过程中
- [ ] 每完成一个小功能就提交一次
- [ ] 提交信息格式参考：
  ```bash
  # 新功能
  git commit -m "Add: 添加xxx功能"

  # 修复问题
  git commit -m "Fix: 修复xxx问题"

  # 优化代码
  git commit -m "Optimize: 优化xxx"

  # 更新文档
  git commit -m "Docs: 更新xxx文档"
  ```

## 结束开发前
- [ ] 检查代码是否都已提交：`git status`
- [ ] 推送到远程：`git push`
- [ ] 检查 README.md 是否需要更新

## Git 命令速查

### 1. 查看状态和历史
bash
查看当前状态
git status
查看提交历史（简洁版）
git log --oneline
查看最近5条历史
git log -5
查看详细修改
git log -p

### 2. 分支操作
bash
查看所有分支
git branch
创建并切换到新分支
git checkout -b feature/xxx
切换回主分支
git checkout main

### 3. 回退操作
bash
查看详细历史
git log -p
回退到指定版本
git reset --hard <commit_id>
查看所有操作记录
git reflog

## 每周五额外检查
- [ ] 完整备份：
  ```bash
  git add .
  git commit -m "Weekly backup: YYYY-MM-DD"
  git push
  ```
- [ ] 检查文档更新
- [ ] 整理本周开发进度