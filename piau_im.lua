-- 設定 Lua Script 模組匯入路徑
package.path = "/Users/alanjui/.luarocks/share/lua/5.4/?.lua;" .. package.path
package.path = "/Users/alanjui/.luarocks/share/lua/5.4/?/init.lua;" .. package.path
package.path = "/Users/alanjui/workspace/rime/piau-im/lua/?.lua;" .. package.path
package.cpath = "/Users/alanjui/.luarocks/lib/lua/5.4/?.so;" .. package.cpath

-- 讀入 .env 設定檔
local env = require("load_env")
if not env then
	print("讀入 .env 設定檔失敗")
	return
end

-- 获取 IMAGE_URL 变量的值
-- local image_url = env.IMAGE_URL
local image_url = env["IMAGE_URL"]
if not image_url then
	-- print("在設定檔中找不到 image_url 變數")
	print(("在設定檔中找不到 %s 變數"):format("image_url"))
end

-- 定義文章標題
-- local title = "洛神賦"
local title = env["TITLE"] or nil
if not title then
	-- print("在設定檔中找不到 title 變數")
	print(("在設定檔中找不到 %s 變數"):format("title"))
end

-- 定義文章標題和注音/拼音方法
local methods = {
	"十五音標音",
	"方音符號注音",
	"白話字拼音",
	"台羅拼音",
	"閩拼拼音",
}

local div_tag = "《%s》【%s】\n"
	.. '<div class="separator" style="clear: both">\n'
	.. '  <a href="圖片" style="display: block; padding: 1em 0; text-align: center">\n'
	.. '    <img alt="" border="0" width="400" data-original-height="630" data-original-width="1200"\n'
	.. '      src="%s" />\n'
	.. "  </a>\n"
	.. "</div>\n"
	.. "\n"

-- 製作每種注音/拼音方法的 HTML Tags
for _, method in ipairs(methods) do
	local output = div_tag:format(title, method, image_url)
	-- vim.api.nvim_echo({ { output } }, false, {})
	print(output)
end
