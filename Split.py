my_string = "world, how and are you?"
my_word = 'world'

# Разбиваем строку на список слов
my_list = my_string.split(my_word)

# Удаляем слово из списка
#my_list.remove(my_word)

# Соединяем список обратно в строку
new_string = "".join(my_list)

print(new_string)