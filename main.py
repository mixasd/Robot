import openpyxl
import pandas as pd



def calc(min_w, max_w,loss_level ):
    d = pd.read_excel("PRC.xlsx")
    big_c = max_w
    small = d.price.rolling(min_w).mean()
    big = d.price.rolling(big_c).mean()

    price_series = d.price.squeeze()
    date_series = d.date_time.squeeze()
    tr_type = 'nan'


    i = big_c - 1
    if small[i] > big[i]:
        tr_type = 'long'
    else:
        tr_type = 'short'

    trade = True

    earinig = 0
#   цена стопа
    stop_price = 0

    price_enter =price_series[i]


    # print('start', date_series[i])
    start_date = date_series[i]

    for ind in range(price_series.size - big_c - 1):
        i = ind + big_c

# стоп для лонга


        # if tr_type == 'long':
        #     if trade:
        #         print(tr_type, price_enter / price_series[i])
        # else:
        #     if trade:
        #         print(tr_type, price_series[i] / price_enter)


        if tr_type == 'long' and price_enter /price_series[i] > loss_level and trade:
            stop_price =  price_series[i]
            trade = False
            # print('!!!!!СТОП')

        if tr_type == 'short' and  price_series[i] / price_enter > loss_level and trade:
            stop_price =  price_series[i]
            trade = False
            # print('!!!!!СТОП')


        if (small[i] > small[i -1])  and (small[i -1] < big[i -1]) and  (small[i] > big[i]):

            if trade:
                stop_price = price_series[i]

            lp = 0
            if tr_type == 'long':
                earinig += stop_price - price_enter
                pl = stop_price - price_enter
            else:
                earinig +=  price_enter - stop_price
                pl = price_enter - stop_price

            # print(tr_type, start_date,date_series[i], price_series[i], price_enter, earinig,pl)

            start_date = date_series[i]
            tr_type = 'long'
            price_enter = price_series[i]
            trade = True


        if (small[i - 1] > small[i ])  and (small[i -1] > big[i -1]) and  (small[i] < big[i]):

            if trade:
                stop_price = price_series[i]

            pl = 0
            if tr_type == 'long':
                earinig += stop_price - price_enter
                pl = stop_price - price_enter
            else:
                earinig +=  price_enter - stop_price
                pl = price_enter - stop_price

            # print(tr_type, start_date, date_series[i], price_series[i], price_enter, earinig,pl)

            start_date = date_series[i]
            tr_type = 'short'
            price_enter = price_series[i]
            trade = True

    return(earinig)


if __name__ == '__main__':

    min_s =0
    min_b = 0
    min_l = 0
    min_e = 10000

    max_s =0
    max_b = 0
    max_l = 0
    max_e = 0


    for s in range(6):
        for b in range(10):
            for l in range(10):
                v = calc(s+2, s+b+2, 1.03+l/50)

                if v > max_e:
                    max_e = v
                    max_s = s+2
                    max_b = s+b+2
                    max_l = 1.03+l/50

                if min_e >v:
                    min_e = v
                    min_s = s+2
                    min_b = s+b+2
                    min_l = 1.03+l/50

    print('min')
    print(min_s)
    print(min_b)
    print(min_l)
    print(min_e)

    print('max')
    print(max_s)
    print(max_b)
    print(max_l)
    print(max_e)

