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

-- Local Function
local function dumpTable(tbl)
	for key, value in pairs(tbl) do
		print(key .. " = " .. tostring(value))
	end
end

-----------------------------------------------------------------
-- 打印所有键值对
dumpTable(env)

-----------------------------------------------------------------
-- 获取 IMAGE_URL 变量的值
print("IMAGE_URL = ", env.IMAGE_URL or "在設定檔中找不到該變數")
