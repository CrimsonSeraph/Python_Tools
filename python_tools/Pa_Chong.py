# 如果环境中未安装 requests，请先执行：
#   pip install requests
import requests
import json
import sys
from typing import Dict, List, Optional, Tuple
import time
import logging

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class WeatherCrawler:
    """天气爬虫类，支持查询多个城市的天气信息"""
    
    def __init__(self):
        """初始化天气爬虫"""
        self.api_key = None  # 用户可以设置API密钥
        self.base_url = "https://api.openweathermap.org/data/2.5/weather"
        self.geo_url = "https://api.openweathermap.org/geo/1.0/direct"
        self.weather_history = []  # 保存查询历史
    
    def set_api_key(self, api_key: str):
        """设置API密钥"""
        self.api_key = api_key
    
    def get_city_coordinates(self, city_name: str) -> Optional[Tuple[float, float]]:
        """获取城市的经纬度坐标"""
        if not self.api_key:
            print("错误: 未设置API密钥")
            return None
        
        try:
            params = {
                'q': city_name,
                'limit': 1,
                'appid': self.api_key
            }
            
            response = requests.get(self.geo_url, params=params, timeout=10)
            response.raise_for_status()
            
            data = response.json()
            if data:
                lat = data[0]['lat']
                lon = data[0]['lon']
                return lat, lon
            else:
                print(f"未找到城市: {city_name}")
                return None
                
        except requests.exceptions.RequestException as e:
            logger.error(f"获取坐标失败: {e}")
            return None
    
    def get_weather_by_coords(self, lat: float, lon: float) -> Optional[Dict]:
        """根据经纬度获取天气信息"""
        if not self.api_key:
            print("错误: 未设置API密钥")
            return None
        
        try:
            params = {
                'lat': lat,
                'lon': lon,
                'appid': self.api_key,
                'units': 'metric',  # 使用摄氏度
                'lang': 'zh_cn'     # 中文
            }
            
            response = requests.get(self.base_url, params=params, timeout=10)
            response.raise_for_status()
            
            return response.json()
            
        except requests.exceptions.RequestException as e:
            logger.error(f"获取天气失败: {e}")
            return None
    
    def get_weather(self, city_name: str) -> Optional[Dict]:
        """获取指定城市的天气信息"""
        print(f"\n正在查询 {city_name} 的天气...")
        
        coords = self.get_city_coordinates(city_name)
        if not coords:
            return None
        
        lat, lon = coords
        weather_data = self.get_weather_by_coords(lat, lon)
        
        if weather_data:
            self.weather_history.append({
                'city': city_name,
                'time': time.strftime('%Y-%m-%d %H:%M:%S'),
                'data': weather_data
            })
        
        return weather_data
    
    def format_weather_info(self, weather_data: Dict, city_name: str) -> str:
        """格式化天气信息输出"""
        try:
            city = weather_data.get('name', city_name)
            temp = weather_data['main']['temp']
            feels_like = weather_data['main']['feels_like']
            humidity = weather_data['main']['humidity']
            pressure = weather_data['main']['pressure']
            wind_speed = weather_data['wind']['speed']
            weather_desc = weather_data['weather'][0]['description']
            
            result = f"\n{'='*50}"
            result += f"\n🌤️  {city} 天气信息"
            result += f"\n{'='*50}"
            result += f"\n📅 查询时间: {time.strftime('%Y-%m-%d %H:%M:%S')}"
            result += f"\n🌡️  温度: {temp}°C (体感: {feels_like}°C)"
            result += f"\n💧  湿度: {humidity}%"
            result += f"\n🌬️  风速: {wind_speed} m/s"
            result += f"\n🌫️  气压: {pressure} hPa"
            result += f"\n🌈  天气: {weather_desc}"
            result += f"\n{'='*50}\n"
            
            return result
            
        except KeyError as e:
            logger.error(f"解析天气数据失败: {e}")
            return "解析天气数据失败"
    
    def display_weather(self, city_name: str):
        """显示城市天气"""
        weather_data = self.get_weather(city_name)
        if weather_data:
            print(self.format_weather_info(weather_data, city_name))
        else:
            print(f"获取 {city_name} 的天气信息失败")
    
    def display_history(self):
        """显示查询历史"""
        if not self.weather_history:
            print("\n暂无查询历史")
            return
        
        print(f"\n{'='*50}")
        print("查询历史")
        print(f"{'='*50}")
        
        for i, record in enumerate(self.weather_history, 1):
            print(f"\n[{i}] {record['city']} - {record['time']}")
            # 显示简要信息
            try:
                temp = record['data']['main']['temp']
                desc = record['data']['weather'][0]['description']
                print(f"    温度: {temp}°C, 天气: {desc}")
            except:
                pass
        
        print(f"\n{'='*50}")
    
    def save_weather_to_file(self, filename: str = "weather_history.json"):
        """保存查询历史到文件"""
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(self.weather_history, f, ensure_ascii=False, indent=2)
            print(f"\n查询历史已保存到: {filename}")
        except Exception as e:
            logger.error(f"保存文件失败: {e}")
            print(f"保存失败: {e}")
    
    def load_weather_from_file(self, filename: str = "weather_history.json"):
        """从文件加载查询历史"""
        try:
            import os
            if os.path.exists(filename):
                with open(filename, 'r', encoding='utf-8') as f:
                    self.weather_history = json.load(f)
                print(f"\n已加载 {len(self.weather_history)} 条历史记录")
            else:
                print(f"\n文件 {filename} 不存在")
        except Exception as e:
            logger.error(f"加载文件失败: {e}")
            print(f"加载失败: {e}")

def get_input(prompt: str) -> str:
    """获取用户输入"""
    while True:
        user_input = input(prompt).strip()
        if user_input:
            return user_input
        print("输入不能为空，请重新输入")

def display_menu():
    """显示主菜单"""
    print("\n" + "="*60)
    print("           🌤️  天气查询工具")
    print("="*60)
    print("1. 查询城市天气")
    print("2. 查看查询历史")
    print("3. 保存查询历史")
    print("4. 加载查询历史")
    print("5. 设置API密钥")
    print("6. 退出程序")
    print("="*60)

def setup_api_key(crawler: WeatherCrawler):
    """设置API密钥"""
    print("\n📝 设置API密钥")
    print("-"*40)
    print("注意: 本程序使用 OpenWeatherMap API")
    print("获取免费API密钥请访问: https://openweathermap.org/api")
    print("-"*40)
    
    api_key = get_input("请输入API密钥: ")
    crawler.set_api_key(api_key)
    print("\n✅ API密钥已设置")

def query_weather(crawler: WeatherCrawler):
    """查询天气"""
    if not crawler.api_key:
        print("\n⚠️  请先设置API密钥 (菜单选项5)")
        return
    
    while True:
        city_name = get_input("\n请输入要查询的城市名称 (输入 'q' 返回): ")
        
        if city_name.lower() == 'q':
            break
        
        crawler.display_weather(city_name)
        
        choice = input("是否继续查询其他城市? (Y/N): ").strip().lower()
        if choice not in ['y', 'yes', '']:
            break

def main():
    """主函数"""
    print("=" * 60)
    print("           🌤️  天气查询工具")
    print("=" * 60)
    
    try:
        # 初始化爬虫
        crawler = WeatherCrawler()
        
        # 主循环
        while True:
            display_menu()
            choice = get_input("请选择操作 (1-6): ")
            
            if choice == '1':
                query_weather(crawler)
            elif choice == '2':
                crawler.display_history()
            elif choice == '3':
                filename = input("请输入保存文件名 (默认: weather_history.json): ").strip()
                if not filename:
                    filename = "weather_history.json"
                crawler.save_weather_to_file(filename)
            elif choice == '4':
                filename = input("请输入加载文件名 (默认: weather_history.json): ").strip()
                if not filename:
                    filename = "weather_history.json"
                crawler.load_weather_from_file(filename)
            elif choice == '5':
                setup_api_key(crawler)
            elif choice == '6':
                print("\n👋 感谢使用天气查询工具！")
                break
            else:
                print("无效的选择，请输入 1-6")
            
            input("\n按 Enter 键继续...")
        
    except KeyboardInterrupt:
        print("\n\n用户中断操作")
    except Exception as e:
        print(f"\n发生错误: {e}")
        import traceback
        traceback.print_exc()
    
    input("\n按 Enter 键退出...")

if __name__ == "__main__":
    main()
