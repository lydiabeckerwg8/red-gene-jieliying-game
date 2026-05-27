# 红色基因接力营

这是一个可以直接托管为网页链接的单文件静态小游戏。入口文件是 `index.html`，不需要安装依赖或后端服务。当前版本已升级为“研学路线图 + 精神火种 + 研学护照 + 主题卡片收集”的地图闯关式互动体验。

## 当前在线链接

已发布到 GitHub Pages，可直接访问：

```text
https://lydiabeckerwg8.github.io/red-gene-jieliying-game/
```

## 在线发布

推荐使用 GitHub Pages：

1. 在 GitHub 新建一个公开仓库，例如 `red-gene-mini-game`。
2. 上传本文件夹内的全部文件，包含 `index.html`、`.nojekyll` 和 `.github/workflows/deploy-pages.yml`。
3. 打开仓库的 `Settings` -> `Pages`。
4. 在 `Build and deployment` 中将 `Source` 选择为 `GitHub Actions`。
5. 回到仓库的 `Actions` 页面，等待 `Deploy to GitHub Pages` 运行完成。

发布成功后，访问链接通常是：

```text
https://你的GitHub用户名.github.io/red-gene-mini-game/
```

如果仓库名不是 `red-gene-mini-game`，把链接最后一段改成你的仓库名。

## 本地预览

直接双击 `index.html` 即可在浏览器中打开。也可以把整个文件夹上传到 Netlify、Vercel、Cloudflare Pages 等静态站点平台，发布目录选择项目根目录即可。
