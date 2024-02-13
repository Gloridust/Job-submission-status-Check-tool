from ltp import LTP
ltp = LTP()
word = ltp.pipeline(['秦卓230110105.doc', '秦悦230110130.docx'], tasks=["cws"], return_dict=False)
word_flat = [item for sublist in word[0] for item in sublist]
print(word_flat)