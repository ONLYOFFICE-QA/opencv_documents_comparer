# from requests_html import HTMLSession
#
#
# session = HTMLSession()
# r = session.get('https://python.org')
#
# pyver = r.html.find('.small-widget.download-widget a', first=True)
#
# print(pyver.text)
#
def dict_compare(d1, d2):
    d1_keys = set(d1.keys())
    d2_keys = set(d2.keys())
    shared_keys = d1_keys.intersection(d2_keys)
    added = d1_keys - d2_keys
    removed = d2_keys - d1_keys
    modified = {o: (f'Befor {d1[o]}', f'After {d2[o]}') for o in shared_keys if d1[o] != d2[o]}
    same = set(o for o in shared_keys if d1[o] == d2[o])
    return added, removed, modified, same


x = dict(
    {"num_of_sheets": "8", "number_of_lines": "277", "word_count": "830", "number_of_characters_without_spaces": "4896",
     "number_of_characters_with_spaces": "5714", "number_of_abzad": "45"})
y = dict({"num_of_sheets": "8", "number_of_lines": "275555", "word_count": "830",
          "number_of_characters_without_spaces": "4896",
          "number_of_characters_with_spaces": "5714", "number_of_abzad": "45"})
added, removed, modified, same = dict_compare(x, y)
print(modified.keys())
print(x['number_of_lines'], y['number_of_lines'])
for i in modified.keys():
    if i == 'number_of_lines':
        if int(x['number_of_lines']) > int(y['number_of_lines']):
            print('ffffff')
        else:
            print('dfsfsdfsdfsdfsdfsdf')
