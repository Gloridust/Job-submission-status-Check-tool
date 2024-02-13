from ltp import LTP
ltp = LTP()
word = ltp.pipeline("他叫汤姆去拿外衣。阙棵2301101024？", tasks=["cws"], return_dict = False)
print(word)