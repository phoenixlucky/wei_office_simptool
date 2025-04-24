import pandas as pd
from statsmodels.tsa.arima.model import ARIMA
from collections import Counter
import jieba
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
from wordcloud import WordCloud
from statsmodels.tsa.stattools import adfuller
from statsmodels.tsa.arima_model import auto_arima
import pmdarima as pm
import warnings

class TrendPredictor:
    def __init__(self, market_trend_df, date_col, smoothed_avg_col,
                 rise_label='上升', fall_label='下滑', flat_label='横盘',
                 freq='B', order=None, steps=7, sortdata='逆序', auto_order=True):
        self.market_trend_df = market_trend_df.copy()
        self.date_col = date_col
        self.smoothed_avg_col = smoothed_avg_col
        self.rise_label = rise_label
        self.fall_label = fall_label
        self.flat_label = flat_label
        self.freq = freq
        self.order = order if order else (5, 1, 0)
        self.steps = steps
        self.sortdata = sortdata
        self.auto_order = auto_order
        self._prepare_data()

    def _prepare_data(self):
        if self.sortdata == '逆序':
            self.reversed_market_trend_df = self.market_trend_df[self.smoothed_avg_col][::-1].reset_index(drop=True)
        else:
            self.reversed_market_trend_df = self.market_trend_df[self.smoothed_avg_col].reset_index(drop=True)

        self.market_trend_df['趋势'] = self.market_trend_df[self.smoothed_avg_col].diff().apply(
            lambda x: self.rise_label if x > 0 else (self.fall_label if x < 0 else self.flat_label))
        
        # 检查数据平稳性
        self.is_stationary = self._check_stationarity(self.reversed_market_trend_df)
        
        # 如果启用自动参数选择，确定最佳ARIMA参数
        if self.auto_order:
            self.order = self._find_best_order()

    def _check_stationarity(self, series):
        """检查时间序列是否平稳"""
        result = adfuller(series.dropna())
        return result[1] <= 0.05  # p值小于0.05表示序列是平稳的
    
    def _find_best_order(self):
        """使用auto_arima自动寻找最佳ARIMA参数"""
        warnings.filterwarnings("ignore")
        # 使用pmdarima的auto_arima函数自动寻找最佳参数
        model = pm.auto_arima(
            self.reversed_market_trend_df,
            start_p=0, start_q=0,
            max_p=5, max_q=5, max_d=2,
            seasonal=False,
            trace=False,
            error_action='ignore',
            suppress_warnings=True,
            stepwise=True
        )
        return model.order

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
        # 使用最佳参数或用户指定参数创建ARIMA模型
        model = ARIMA(self.reversed_market_trend_df, order=self.order)
        model_fit = model.fit()
        
        # 获取预测结果及置信区间
        forecast_result = model_fit.forecast(steps=self.steps, alpha=0.05)
        forecast = forecast_result.tolist() if isinstance(forecast_result, np.ndarray) else forecast_result
        forecast = [round(x, 4) for x in forecast]

        last_value = self.market_trend_df[self.smoothed_avg_col][
            self.market_trend_df[self.date_col] == self.market_trend_df[self.date_col].max()].tolist()[0]
        forecast.insert(0, last_value)

        future_dates = pd.date_range(start=self.market_trend_df[self.date_col].max(), periods=len(forecast),
                                     freq=self.freq)
        future_forecast_df = pd.DataFrame({self.date_col: future_dates.date, '预测值': forecast})
        future_forecast_df['趋势'] = future_forecast_df['预测值'].diff().apply(
            lambda x: self.rise_label if x > 0 else (self.fall_label if x < 0 else self.flat_label))

        # 添加模型评估指标
        if len(self.reversed_market_trend_df) > 10:
            self.model_metrics = {
                'AIC': round(model_fit.aic, 2),
                'BIC': round(model_fit.bic, 2),
                'RMSE': round(np.sqrt(model_fit.mse), 4)
            }
        else:
            self.model_metrics = {'注意': '数据量不足，无法计算可靠的模型评估指标'}

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
    
    def get_model_info(self):
        """返回模型信息和评估指标"""
        info = {
            '模型参数': f"ARIMA{self.order}",
            '数据是否平稳': "是" if self.is_stationary else "否",
            '预测步数': self.steps
        }
        
        # 如果已经进行过预测，添加模型评估指标
        if hasattr(self, 'model_metrics'):
            info.update(self.model_metrics)
            
        return info
    
    def cross_validate(self, test_size=0.2):
        """使用时间序列交叉验证评估模型性能"""
        if len(self.reversed_market_trend_df) < 10:
            return {'错误': '数据量不足，无法进行交叉验证'}
            
        # 划分训练集和测试集
        train_size = int(len(self.reversed_market_trend_df) * (1 - test_size))
        train, test = self.reversed_market_trend_df[:train_size], self.reversed_market_trend_df[train_size:]
        
        # 使用训练集拟合模型
        model = ARIMA(train, order=self.order)
        model_fit = model.fit()
        
        # 预测测试集
        predictions = model_fit.forecast(steps=len(test))
        
        # 计算评估指标
        mse = np.mean((predictions - test) ** 2)
        rmse = np.sqrt(mse)
        mae = np.mean(np.abs(predictions - test))
        mape = np.mean(np.abs((test - predictions) / test)) * 100
        
        return {
            '均方误差(MSE)': round(mse, 4),
            '均方根误差(RMSE)': round(rmse, 4),
            '平均绝对误差(MAE)': round(mae, 4),
            '平均绝对百分比误差(MAPE)': round(mape, 2)
        }
        
# Example usage:
# Assuming market_trend_df is a DataFrame with columns '日期' and '平滑平均'
# predictor = TrendPredictor(market_trend_df, '日期', '平滑平均', rise_label='上升', fall_label='下滑', flat_label='横盘',sortdata='逆序')
# original_df = predictor.original_data()
# future_forecast_df, forecast, str_forecast, future_dates = predictor.forecast_data()
# future7_df, forecast, str_forecast, future_dates = predictor.styled_forecast_data()
class MultipleTrendPredictor():
    def __init__(self, market_trend_df, freq='B', order=(5, 1, 0), steps=7):
        self.market_trend_df = market_trend_df.copy()
        self.freq = freq
        self.order = order
        self.steps = steps

    def predict(self):
        # 按索引的时间顺序排序
        self.market_trend_df = self.market_trend_df.sort_index(ascending=True)

        # 预测函数
        def predict_next_days(series, days):
            model = ARIMA(series, order=self.order)
            model_fit = model.fit()
            forecast = model_fit.forecast(steps=days)
            return forecast

        # 预测
        predictions = pd.DataFrame()
        for column in self.market_trend_df.columns:
            predictions[column] = predict_next_days(self.market_trend_df[column].reset_index(drop=True), self.steps)

        # 创建预测结果数据框
        last_date = self.market_trend_df.index.max()+ pd.Timedelta(days=1)
        future_dates = pd.date_range(start=last_date, freq=self.freq, periods=self.steps)
        predictions.index = future_dates
        return predictions

class TextAnalysis:
    def __init__(self, dataframe):
        self.df = dataframe

    def get_word_freq(self, group_col, text_col, agg_func):
        # 聚合数据
        aggregated_text = self.df.groupby(group_col)[text_col].apply(agg_func).reset_index()
        # 计算词频
        aggregated_text['word_freq'] = aggregated_text[text_col].apply(self.compute_word_freq)
        return aggregated_text

    def compute_word_freq(self, text):
        words = jieba.cut(text)
        return Counter(words)

    def plot_wordclouds(self, word_freqs, titles, save_path="wordclouds.png"):
        def create_ellipse_mask(width, height):
            y, x = np.ogrid[-height // 2:height // 2, -width // 2:width // 2]
            mask = (x ** 2 / (width // 2) ** 2 + y ** 2 / (height // 2) ** 2) <= 1
            mask = 255 * mask.astype(int)
            return mask

        ellipse_mask = create_ellipse_mask(400, 200)
        ellipse_mask = 255 - ellipse_mask  # 反转掩码

        num_plots = len(word_freqs)
        cols = 2
        rows = (num_plots + 1) // cols

        fig, axes = plt.subplots(rows, cols, figsize=(16, 8))

        for i, (word_freq, title) in enumerate(zip(word_freqs, titles)):
            ax = axes[i // cols, i % cols]
            wordcloud = WordCloud(
                width=400,
                height=200,
                max_words=200,
                font_path='C:/Windows/Fonts/SimHei.ttf',
                background_color='white',
                mask=ellipse_mask
            ).generate_from_frequencies(word_freq)

            ax.imshow(wordcloud, interpolation='bilinear')
            ax.set_title(title)
            ax.axis('off')
            ax.set_xticks([])  # 添加这行代码
            ax.set_yticks([])  # 添加这行代码

        for j in range(i + 1, rows * cols):
            fig.delaxes(axes[j // cols, j % cols])

        plt.axis('off')  # 添加这行代码
        plt.tight_layout()
        plt.savefig(save_path, bbox_inches='tight')  # 保存图像
        plt.close()  # 关闭图像以释放内存

# 使用示例
# dataframe = ...  # 假设你已经有了一个 pandas DataFrame
# text_analysis = TextAnalysis(dataframe)
# word_freqs, titles = ...  # 假设你已经准备好了词频和标题
# text_analysis.plot_wordclouds(word_freqs, titles, save_path="wordclouds.png")
# # 创建预测器实例，启用自动参数选择
# predictor = TrendPredictor(market_trend_df, '日期', '平滑平均', auto_order=True)

# # 获取预测结果
# future_forecast_df, forecast, str_forecast, future_dates = predictor.forecast_data()

# # 获取模型信息和评估指标
# model_info = predictor.get_model_info()
# print(model_info)

# # 进行交叉验证
# cv_results = predictor.cross_validate(test_size=0.2)
# print(cv_results)