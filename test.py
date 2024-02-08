# import re

# title = "[27!가a]안녕"

# # 대괄호 안의 숫자와 문자를 찾는 정규표현식
# pattern = re.compile(r'\[([^]]*\d[^]]*)\]')

# # 대괄호 안의 숫자와 문자를 찾아서 대괄호와 숫자만 남기고 내용을 제거하는 함수
# def process_title(match):
#     content = match.group(1)
#     numbers = re.findall(r'\d+', content)
#     if numbers:
#         return '[' + numbers[0] + ']'
#     else:
#         return ''

# # 정규표현식 패턴을 사용하여 문자열 처리
# title = pattern.sub(process_title, title).strip()

# print(title)

import re

title = "[가나다]안녕"
pattern = re.compile(r'\[[^\]\d]*\]')

title = re.sub(pattern, '',title)

print(title)