import pandas as pd
from statsmodels.tsa.arima.model import ARIMA


class TrendPredictor:
    def __init__(self, market_trend_df, date_col, smoothed_avg_col,
                 rise_label='上升', fall_label='下滑', flat_label='横盘',
                 freq='B', order=(5, 1, 0), steps=7, sortdata='逆序'):
        self.market_trend_df = market_trend_df.copy()
        self.date_col = date_col
        self.smoothed_avg_col = smoothed_avg_col
        self.rise_label = rise_label
        self.fall_label = fall_label
        self.flat_label = flat_label
        self.freq = freq
        self.order = order
        self.steps = steps
        self.sortdata = sortdata
        self._prepare_data()

    def _prepare_data(self):
        if self.sortdata == '逆序':
            self.reversed_market_trend_df = self.market_trend_df[self.smoothed_avg_col][::-1].reset_index(drop=True)
        else:
            self.reversed_market_trend_df = self.market_trend_df[self.smoothed_avg_col].reset_index(drop=True)

        self.market_trend_df['趋势'] = self.market_trend_df[self.smoothed_avg_col].diff().apply(
            lambda x: self.rise_label if x > 0 else (self.fall_label if x < 0 else self.flat_label))

    def original_data(self):
        return self.market_trend_df

    def _highlight_color(self, val):
        if val == self.rise_label:
            color = 'crimson'
        elif val == self.fall_label:
            color = 'forestGreen'
        else:
            color = 'black'
        return f'color: {color}'

    def _predict(self):
        model = ARIMA(self.reversed_market_trend_df, order=self.order)
        model_fit = model.fit()
        forecast = model_fit.forecast(steps=self.steps).tolist()
        forecast = [round(x, 4) for x in forecast]

        last_value = self.market_trend_df[self.smoothed_avg_col][
            self.market_trend_df[self.date_col] == self.market_trend_df[self.date_col].max()].tolist()[0]
        forecast.insert(0, last_value)

        future_dates = pd.date_range(start=self.market_trend_df[self.date_col].max(), periods=len(forecast),
                                     freq=self.freq)
        future_forecast_df = pd.DataFrame({self.date_col: future_dates.date, '预测值': forecast})
        future_forecast_df['趋势'] = future_forecast_df['预测值'].diff().apply(
            lambda x: self.rise_label if x > 0 else (self.fall_label if x < 0 else self.flat_label))

        future_forecast_df = pd.DataFrame(
            future_forecast_df[future_forecast_df[self.date_col] > self.market_trend_df[self.date_col].max()],
            columns=[self.date_col, "预测值", '趋势'])

        return future_forecast_df, forecast, list(map(str, forecast)), future_dates

    def forecast_data(self):
        future_forecast_df, forecast, str_forecast, future_dates = self._predict()
        return future_forecast_df, forecast, str_forecast, future_dates

    def styled_forecast_data(self):
        future_forecast_df, forecast, str_forecast, future_dates = self._predict()
        future_forecast_df['预测值'] = future_forecast_df['预测值'].astype(str)
        future_forecast_df = future_forecast_df.set_index(self.date_col).T
        future7_df = future_forecast_df.style.map(
            lambda val: self._highlight_color(val),
            subset=pd.IndexSlice['趋势', :])

        return future7_df, forecast, str_forecast, future_dates

# Example usage:
# Assuming market_trend_df is a DataFrame with columns '日期' and '平滑平均'
# predictor = TrendPredictor(market_trend_df, '日期', '平滑平均', rise_label='上升', fall_label='下滑', flat_label='横盘',sortdata='逆序')
# original_df = predictor.original_data()
# future_forecast_df, forecast, str_forecast, future_dates = predictor.forecast_data()
# future7_df, forecast, str_forecast, future_dates = predictor.styled_forecast_data()

