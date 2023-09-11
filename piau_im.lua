-- 定義文章標題和注音/拼音方法
local title = "歸去來辭"
local methods = {
	"十五音標音",
	"方音符號注音",
	"白話字拼音",
	"台羅拼音",
	"閩拼拼音",
}

-- 定義圖片的 URL
local img_url = "https://p3-tt.byteimg.com/origin/pgc-image/92b3838d44574423aa662573d30043ef?from=pc"

local format = "《%s》【%s】\n"
	.. '<div class="separator" style="clear: both">\n'
	.. '  <a href="圖片" style="display: block; padding: 1em 0; text-align: center">\n'
	.. '    <img alt="" border="0" width="400" data-original-height="630" data-original-width="1200"\n'
	.. '      src="%s" />\n'
	.. "  </a>\n"
	.. "</div>\n"
	.. "\n"

-- 製作每種注音/拼音方法的 HTML Tags
for _, method in ipairs(methods) do
	local output = format:format(title, method, img_url)
	vim.api.nvim_echo({ { output } }, false, {})
end
