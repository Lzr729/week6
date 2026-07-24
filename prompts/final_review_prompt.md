# Final人工复核指令

对复核队列逐条确认：
- 接受当前Auto结果时，before_value与after_value相同；
- 重复或错误主体被排除时，after_value为null；
- 披露限制记录不补写未披露事实；
- 每条决定保留原因、证据ID、PDF页码和正文页码。
