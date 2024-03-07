-- my-dotenv.lua

local M = {}

function M.get(variableName)
	-- 打开 .env 文件
	local f = io.open(".env", "r")

	if not f then
		return nil, ".env 文件不存在或无法打开"
	end

	-- 初始化变量
	local variableValue = nil

	-- 遍历文件的每一行
	for line in f:lines() do
		-- 使用正则表达式匹配以指定变量名称开头的行
		local key, value = line:match("^" .. variableName .. "=(.*)$")
		vim.api.nvim_echo({ { key, "Normal" } }, false, {})
		vim.api.nvim_echo({ { value, "Normal" } }, false, {})

		if key and value then
			variableValue = value
			break
		end
	end

	-- 关闭文件
	f:close()

	-- 返回变量的值或错误消息
	if variableValue then
		return variableValue
	else
		-- return nil, ("未找到变量 %s"):format(variableName)
		local err_msg = string.format("未找到变量 %s", variableName)
		return nil, err_msg
	end
end

return M
