import pickle

from rtneo_address import WeHave
from searcher import Search

a = WeHave('зима', 'Зима', '38:35')

a.create_rtneo_file()

a = Search('')

a.put_daughter()

a.put_info(5000)

a.reformat_uk_json()
a.put_uk_info()

with open('data.pickle', 'wb') as f:
    pickle.dump(a, f)
