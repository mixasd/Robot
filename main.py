import numpy as np
import openpyxl
import pandas as pd


def calc_2(min_w, max_w, loss_level):
    pd.set_option('display.max_rows', 1000)
    pd.set_option('display.max_columns', 100)

    rsi_p = 14
    rsi_oversold = 30
    rsi_overbought = 70

    data = pd.read_excel("PRC.xlsx")
    data['ma_short'] = data['price'].rolling(min_w).mean()
    data['ma_long'] = data['price'].rolling(max_w).mean()
    data['ma_signal'] = data['ma_short'] > data['ma_long']
    data['up'] = np.maximum(data['price'].diff(), 0)
    data['down'] = np.maximum(-data['price'].diff(), 0)
    data['RS'] = data['up'].rolling(rsi_p).mean() / data['down'].rolling(rsi_p).mean()
    data['RSI'] = 100 - 100 / (1 + data['RS'])
    data['RSI_signal'] = 1 * (data['RSI'] < rsi_oversold) - 1 * (data['RSI'] > rsi_overbought)

    # data.to_excel('saved.xlsx', index=False)

    price_enter = data['price'][max_w - 1]
    date_enter = data['date_time'][max_w - 1]
    trade = True
    earinig = 0

    for ind in range(max_w, len(data)):

        if (data['ma_signal'][ind] == data['ma_signal'][ind - 1]) and data['ma_signal'][ind] and price_enter / \
                data['price'][ind] > loss_level and trade:
            stop_price = data['price'][ind]
            trade = False
            # print('!!!!!СТОП',data['date_time'][ind], data['price'][ind] , price_enter,data['ma_signal'][ind] ,date_enter  )

        if (data['ma_signal'][ind] == data['ma_signal'][ind - 1]) and not (data['ma_signal'][ind]) and data['price'][
            ind] / price_enter > loss_level and trade:
            stop_price = data['price'][ind]
            trade = False
            # print('!!!!!СТОП',data['date_time'][ind], data['price'][ind] , price_enter,data['ma_signal'][ind],date_enter   )

        if not (data['ma_signal'][ind - 1]) and data['ma_signal'][ind]:
            if trade:
                stop_price = data['price'][ind]

            lp = 0
            if data['ma_signal'][ind - 1]:
                earinig += stop_price - price_enter
                pl = stop_price - price_enter
            else:
                earinig += price_enter - stop_price
                pl = price_enter - stop_price

            # print(date_enter, data['date_time'][ind],data['ma_signal'][ind],price_enter, data['price'][ind], earinig,pl )
            price_enter = data['price'][ind]
            trade = True
            date_enter = data['date_time'][ind]

        if not (data['ma_signal'][ind]) and data['ma_signal'][ind - 1]:

            if trade:
                stop_price = data['price'][ind]

            pl = 0
            if data['ma_signal'][ind - 1]:
                earinig += stop_price - price_enter
                pl = stop_price - price_enter
            else:
                earinig += price_enter - stop_price
                pl = price_enter - stop_price

            # print(date_enter, data['date_time'][ind], data['ma_signal'][ind], price_enter, data['price'][ind], earinig,pl)
            price_enter = data['price'][ind]
            trade = True
            data['date_time'][ind]

    return (earinig)


def calc(min_w, max_w, loss_level):
    d = pd.read_excel("PRC.xlsx")
    big_c = max_w
    small = d.price.rolling(min_w).mean()
    big = d.price.rolling(big_c).mean()

    price_series = d.price.squeeze()
    date_series = d.date_time.squeeze()
    tr_type = 'nan'

    i = big_c - 1
    if small[i] <= big[i]:
        tr_type = 'short'
    else:
        tr_type = 'long'

    trade = True

    earinig = 0
    #   цена стопа
    stop_price = 0

    price_enter = price_series[i]

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

        if tr_type == 'long' and price_enter / price_series[i] > loss_level and trade:
            stop_price = price_series[i]
            trade = False
            # print('!!!!!СТОП')

        if tr_type == 'short' and price_series[i] / price_enter > loss_level and trade:
            stop_price = price_series[i]
            trade = False
            # print('!!!!!СТОП')

        if (small[i] > small[i - 1]) and (small[i - 1] < big[i - 1]) and (small[i] > big[i]):

            if trade:
                stop_price = price_series[i]

            lp = 0
            if tr_type == 'long':
                earinig += stop_price - price_enter
                pl = stop_price - price_enter
            else:
                earinig += price_enter - stop_price
                pl = price_enter - stop_price

            # print(tr_type, start_date,date_series[i], price_series[i], price_enter, earinig,pl)

            start_date = date_series[i]
            tr_type = 'long'
            price_enter = price_series[i]
            trade = True

        if (small[i - 1] > small[i]) and (small[i - 1] > big[i - 1]) and (small[i] < big[i]):

            if trade:
                stop_price = price_series[i]

            pl = 0
            if tr_type == 'long':
                earinig += stop_price - price_enter
                pl = stop_price - price_enter
            else:
                earinig += price_enter - stop_price
                pl = price_enter - stop_price

            # print(tr_type, start_date, date_series[i], price_series[i], price_enter, earinig,pl)

            start_date = date_series[i]
            tr_type = 'short'
            price_enter = price_series[i]
            trade = True

    return (earinig)


if __name__ == '__main__':
    # print(calc_2(7, 16, 1.11))
    # print(calc_2(7, 16, 1.15))
    # print(calc_2(4, 6, 1.21))
    # print(calc_2(7, 14, 1.05))

    min_s = 0
    min_b = 0
    min_l = 0
    min_e = 10000

    max_s = 0
    max_b = 0
    max_l = 0
    max_e = 0

    for s in range(6):
        for b in range(10):
            for l in range(10):
                v = calc_2(s + 2, s + b + 2, 1.03 + l / 50)
                print(v, s + 2, s + b + 2, 1.03 + l / 50)
                if v > max_e:
                    max_e = v
                    max_s = s + 2
                    max_b = s + b + 2
                    max_l = 1.03 + l / 50

                if min_e > v:
                    min_e = v
                    min_s = s + 2
                    min_b = s + b + 2
                    min_l = 1.03 + l / 50

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
