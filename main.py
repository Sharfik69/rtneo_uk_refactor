import time

from divider import Divider
from rtneo_address import WeHave
from searcher import Search

start_time = time.time()
a = WeHave('зима', 'Зима', '38:35')

a.create_rtneo_file()

a = Search('')

a.put_daughter()

a.put_info(5000)

a.reformat_uk_json()
a.put_uk_info()

b = Divider()
b.divide_by_assignation_code()
b.divide_by_type_uk()

print("%s seconds" % (time.time() - start_time))
