using BlazorWebViewOnWPF.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlazorWebViewOnWPF.Services
{
    class WeatherDataService
    {
        public async Task<IEnumerable<WeatherForecast>> GetWeatherForecastsAsync()
        {
            await Task.Delay(500);
            var stringData = await File.ReadAllTextAsync("weather.json");
            return JsonConvert.DeserializeObject<IEnumerable<WeatherForecast>>(stringData);
        }
    }
}
