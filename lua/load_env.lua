-- 打开 .env 文件
local f = io.open(".env", "r")
if not f then
	return nil, ".env 文件不存在或无法打开"
end

-- 初始化环境变量表
local envTable = {}

-- 遍历文件的每一行
for line in f:lines() do
	local key, value = line:match("^([%a_-]+)=(.*)$")
	if key and value then
		envTable[key] = value
	end
end

-- 关闭文件
f:close()

-- 返回环境变量表
return envTable
