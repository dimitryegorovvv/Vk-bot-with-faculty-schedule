import vk_api
import openpyxl
from vk_api.longpoll import VkLongPoll, VkEventType
from vk_api.keyboard import VkKeyboard, VkKeyboardColor
vk_session = vk_api.VkApi(token = "")
session_api = vk_session.get_api()
longpoll = VkLongPoll(vk_session)

current_menu_is = 'none'
faculty_is = 'none'
course_is = 'none'
group_is = 'none'
day_is = 'none'

correct_group_KF='none'
correct_group_PM='none'
correct_group_EK='none'
correct_group_MI='none'
correct_group_FI='none'
mi_kn='none'

correct_excel = openpyxl.open('correct_schedule.xlsx')
def sender(id, text, keyboard=None):
    post = {
        'user_id' : id,
        'message' : text,
        'random_id' : 0
    }
    if keyboard != None:
        post["keyboard"] = keyboard.get_keyboard()
    else:
        post = post
    vk_session.method('messages.send', post)

for event in longpoll.listen():
    if event.type == VkEventType.MESSAGE_NEW and event.to_me:
            msg = event.text.lower()
            id = event.user_id
            if msg == 'назад' or msg=='начать' or msg=='start':
                current_menu_is = 'choise_of_faculty'
                keyboard = VkKeyboard()
                keyboard.add_button('физ.мат')
                sender(id, 'выберите факультет', keyboard)
            elif (msg == 'физ.мат'):
                current_menu_is='choise_of_course'
                faculty_is = msg
                keyboard = VkKeyboard()
                keyboard.add_button('1')
                keyboard.add_button('2')
                keyboard.add_button('3')
                keyboard.add_line()
                keyboard.add_button('4')
                keyboard.add_button('маг-1')
                keyboard.add_button('маг-2')
                keyboard.add_line()
                keyboard.add_button('назад', VkKeyboardColor.NEGATIVE)
                sender(id, 'выберите курс', keyboard)
            # физ.мат
            elif ((msg == '1' or msg == '2' or msg == '3' or msg =='4' or msg == 'маг-1'
                   or msg == 'маг-2') and faculty_is == 'физ.мат' and
                  current_menu_is=='choise_of_course'):
                current_menu_is='choise_of_group'
                course_is = msg
                if (msg == 'маг-1' or msg == 'маг-2'):
                    correct_sheet = correct_excel['МАГ 1-2 КУРС']
                    mi_kn = 'МиКН (маг) - 22'
                    keyboard = VkKeyboard()
                    keyboard.add_button(mi_kn)
                    keyboard.add_line()
                    keyboard.add_button('назад', VkKeyboardColor.NEGATIVE)
                    sender(id, 'выберите группу', keyboard)
                else:
                    correct_sheet = correct_excel.worksheets[int(course_is) - 1]
                    kf = 'кф-x1'
                    correct_group_KF = kf[:3] + msg + kf[4:]
                    pm = 'пм-x1'
                    correct_group_PM = pm[:3] + msg + pm[4:]
                    ek = 'эк-x1'
                    correct_group_EK = ek[:3] + msg + ek[4:]
                    mi = 'ми-x1'
                    correct_group_MI = mi[:3] + msg + mi[4:]
                    fi = 'фи-x1'
                    correct_group_FI = fi[:3] + msg + fi[4:]
                    keyboard = VkKeyboard()
                    keyboard.add_button(correct_group_KF)
                    keyboard.add_button(correct_group_PM)
                    keyboard.add_button(correct_group_EK)
                    keyboard.add_line()
                    keyboard.add_button(correct_group_MI)
                    keyboard.add_button(correct_group_FI)
                    keyboard.add_line()
                    keyboard.add_button('назад', VkKeyboardColor.NEGATIVE)
                    sender(id, 'выберите группу', keyboard)

            elif ((msg == correct_group_KF or msg == correct_group_PM or msg == correct_group_EK
                or msg == correct_group_MI or msg == correct_group_FI or msg == mi_kn.lower())
                and current_menu_is=='choise_of_group'):
                current_menu_is='choise_of_day'
                if msg == mi_kn.lower():
                    msg='МиКН (маг) - 22'
                group_is = msg
                ########################excel#########################
                group_found = False
                max_cell_is = 'Z' + str(correct_sheet.max_row)
                cell_range = correct_sheet['A1':max_cell_is]
                for cell in cell_range:
                    if group_found:
                        break
                    for row_with_group_is in cell:
                        if group_found:
                            break
                        if msg == mi_kn:
                            if row_with_group_is.value == group_is:
                                row_with_group_is = int(row_with_group_is.col_idx)
                                group_found = True
                                break
                        else:
                            if row_with_group_is.value == group_is.upper():
                                row_with_group_is = int(row_with_group_is.col_idx)
                                group_found = True
                                break
                ########################excel#########################

                if group_found == False:
                    sender(id, 'данная группа не найдена')
                else:
                    keyboard = VkKeyboard()
                    keyboard.add_button('пн')
                    keyboard.add_button('вт')
                    keyboard.add_button('ср')
                    keyboard.add_line()
                    keyboard.add_button('чт')
                    keyboard.add_button('пт')
                    keyboard.add_button('сб')
                    keyboard.add_line()
                    keyboard.add_button('назад', VkKeyboardColor.NEGATIVE)
                    sender(id, 'выберите день недели', keyboard)

            elif ((msg == 'пн' or msg == 'вт' or msg == 'ср' or msg == 'чт' or msg == 'пт'
                  or msg == 'сб') and current_menu_is=='choise_of_day'):
                day_is = msg

                ########################excel#########################
                days_of_week = {'пн': 0, 'вт': 5, 'ср': 10, 'чт': 15, 'пт': 20, 'сб': 25}

                whole_str = ''
                coupe_is = 1
                code_of_coupe = 'x&#8419;'
                for row in correct_sheet.iter_rows(min_row=6 + int(days_of_week[day_is]),
                                                   max_row=9 + int(days_of_week[day_is]),
                                                   min_col=row_with_group_is,
                                                   max_col=row_with_group_is):
                    for cell in row:
                        smile_code=(code_of_coupe[:0]+
                                               str(coupe_is)+code_of_coupe[1:])
                        if cell.value == None:
                            whole_str = (whole_str + smile_code + '\n'
                                         + 'Нет пары' + '\n')
                            coupe_is = coupe_is + 1
                            continue
                        str1 = cell.value
                        str2 = str1.split()
                        str3 = ' '.join(str2)
                        whole_str = whole_str + smile_code + '\n' + str3 + '\n'
                        coupe_is=coupe_is+1
                ########################excel#########################

                sender(id, whole_str)
            else:
                sender(id, 'команда не распознана, следуйте подсказкам ниже')
