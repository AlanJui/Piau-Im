-- 打开 .env 文件
local f = io.open(".env", "r")

if not f then
	return nil, ".env 文件不存在或无法打开"
end

-- 初始化变量
local variableName = "IMAGE_URL"
local variableValue = nil

-- 遍历文件的每一行
for line in f:lines() do
	-- 使用正则表达式匹配以指定变量名称开头的行
	-- local key, value = line:match("^" .. variableName .. "=(.*)$")
	-- local key, value = line:match(("^(%w+)=(.*)$"):format(variableName))
	-- print("line:", line)
	-- print(string.match(line, "^(%w+)=(.*)$"))
	-- print(string.match(line, "^%w+=(.*)$"))
	-- print(string.match(line, "=(.*)$"))
	local match_value = string.match(line, "=(.*)$")
	print("match_value:", match_value)
end

-- 关闭文件
f:close()
