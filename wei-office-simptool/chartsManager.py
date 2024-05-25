import pandas as pd
from statsmodels.tsa.arima.model import ARIMA


def highlight_color(val, rise_label, fall_label):
    if val == rise_label:
        color = 'crimson'
    elif val == fall_label:
        color = 'forestGreen'
    else:
        color = 'black'
    return f'color: {color}'

def prediction_function(market_trend_df, date_col, smoothed_avg_col,
                        rise_label='上升', fall_label='下滑', flat_label='横盘',
                        freq='B', order=(5, 1, 0), steps=7,sortdata="逆序"):
    if sortdata=='逆序':
        reversed_market_trend_df = market_trend_df[smoothed_avg_col][::-1].reset_index(drop=True)
        market_trend_df['趋势'] = market_trend_df[smoothed_avg_col].diff().apply(
            lambda x: rise_label if x > 0 else (fall_label if x < 0 else flat_label))
    else:
        reversed_market_trend_df = market_trend_df[smoothed_avg_col].reset_index(drop=True)
        market_trend_df['趋势'] = market_trend_df[smoothed_avg_col].diff().apply(
            lambda x: rise_label if x > 0 else (fall_label if x < 0 else flat_label))
    model = ARIMA(reversed_market_trend_df, order=order)
    model_fit = model.fit()
    forecast = model_fit.forecast(steps=steps).tolist()
    forecast = [round(x, 4) for x in forecast]

    last_value =market_trend_df[smoothed_avg_col][market_trend_df[date_col] == market_trend_df[date_col].max()].tolist()[0]
    forecast.insert(0, last_value)

    future_dates = pd.date_range(start=market_trend_df[date_col].max(), periods=len(forecast), freq=freq)
    future_forecast_df = pd.DataFrame({date_col: future_dates.date, '预测值': forecast})
    future_forecast_df['趋势'] = future_forecast_df['预测值'].diff().apply(
        lambda x: rise_label if x > 0 else (fall_label if x < 0 else flat_label))
    future_forecast_df = pd.DataFrame(
        future_forecast_df[future_forecast_df[date_col] > market_trend_df[date_col].max()],
        columns=[date_col, "预测值", '趋势'])
    future_forecast_df['预测值'] = future_forecast_df['预测值'].astype(str)
    future_forecast_df = future_forecast_df.set_index(date_col).T
    future7_df = future_forecast_df.style.applymap(lambda val: highlight_color(val, rise_label, fall_label),
                                                   subset=pd.IndexSlice['趋势', :])

    return future7_df, forecast, list(map(str, forecast)), future_dates

# Example usage:
# future7_df, forecast, str_forecast, future_dates = prediction_function(market_trend_df, '日期', '平滑平均', rise_label='上升', fall_label='下滑', flat_label='横盘')
