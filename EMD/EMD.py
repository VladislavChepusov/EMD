# Работа с экселем (можно заменить на pandas)
import openpyxl as xl 
import numpy as np
import matplotlib.pyplot as plt
import emd 


# Получение  локальных экстремумов
#idx_minimas = argrelextrema(curs, np.less_equal,order=1)[0]
#idx_maximas = argrelextrema(curs, np.greater)[0]
#peak_locs, peak_mags = emd.sift.get_padded_extrema(curs, pad_width=0, mode='peaks')
#trough_locs, trough_mags = emd.sift.get_padded_extrema(curs, pad_width=1, mode='troughs')


# нахождение экстремумов/кубический сплайн/Находит функции средних значения,а также приближения 
def approximation(x):
    proto_imf = x.copy()
    upper_env = emd.sift.interp_envelope(proto_imf, mode='upper',interp_method ='pchip')
    lower_env = emd.sift.interp_envelope(proto_imf, mode='lower',interp_method ='pchip')
    if (upper_env is None or lower_env is None):
        avg_env = upper_env 
        hi = proto_imf
    else:
        avg_env = (upper_env + lower_env) / 2 
        hi = proto_imf - avg_env
    return upper_env,lower_env,avg_env,hi


# Внутренняя итерация 
def Internal_iteration(_curs):
    check = True
    mass_hi_all = [_curs]
    i = 0
   # Для остановки 
    stop = 0.007
    while check:
        i += 1
        _max,_min,_avg,_curs = approximation(_curs)
        mass_hi_all.append(_curs)
        sigma = sum(abs(mass_hi_all[i-1] - mass_hi_all[i])) /sum(mass_hi_all[i-1]**2)
        if (sigma < stop or np.isnan(sigma)):
            check = False
    return mass_hi_all[-1]
    

# Внешняя итерация
def External_iteration(_curs):
    check2 = True
    ri = _curs.copy()
    mass_ci_all = [] 
    mass_ri_all =[]
    j = 0
    while check2:
        j += 1
        ci = Internal_iteration(ri)
        ri = ri - ci 
        if sum(ci) == 0 or sum(ri) == 0:
            check2 = False
        else:
            mass_ci_all.append(ci)
            mass_ri_all.append(ri)
    return  mass_ci_all,mass_ri_all[-1]    
       

# Графическое представление данных 
def graph(datas,curs,new_mass,new_mass2):
    mng = plt.get_current_fig_manager()
    mng.resize(*mng.window.maxsize())  
    size = len(new_mass) + 2
    ax = {}
    for i in range(1,size+1):
        ax[i] = plt.subplot(size+1,1,i)
        ax[i].grid()
    ax[1].title.set_text(f'Динамика курса Евро\Рубль')
    ax[1].plot(datas,curs,color = 'black',label='Данные') 
    ax[1].tick_params(axis='both', which='major', labelsize=6)
    #ax[1].xaxis.set_major_formatter (matplotlib.dates.DateFormatter("%d.%m"))
    #Реконструированный график
    ax[1].plot(sum(new_mass)+new_mass2  ,'r--',label='Реконструкция')
    ax[1].legend(loc="best",prop={'size': 6})
    plt.subplots_adjust(hspace=1) 
    for i in range(2,size):
         ax[i].title.set_text(f'C{i-1}')
         ax[i].plot(datas,new_mass[i-2],color = 'green') 
         ax[i].tick_params(axis='both', which='major', labelsize=6)
    ax[i+1].title.set_text(f'R{i-1}')
    ax[i+1].plot(datas,new_mass2,color = 'blue')  
    ax[i+1].tick_params(axis='both', which='major', labelsize=6)
    plt.show()
    #plt.savefig(f'END/curs_evro.png')
        

def main():
    # Чтение экселя
    book = xl.open('EURO.xlsx',read_only=True)
    sheet =book.active
    data = []
    curs = []
    # Считывание данных и занесение в список
    for row in range (1,sheet.max_row):
        data.append(sheet[row +1][1].value.strftime("%d.%m."))
        curs.append(sheet[row + 1][2].value)
    curs  = np.array(curs)
    ci_mass,rn = External_iteration(curs)
    graph(data,curs,ci_mass,rn)


if __name__ == '__main__':
    main()