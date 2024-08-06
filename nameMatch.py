from thefuzz import fuzz, process

# Strings to compare
s1 = 'Ada County'
s2 = 'ADA'


#FUZZ
print(fuzz.token_set_ratio(s1,s2))
