using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using System;
using System.IO;
using Newtonsoft.Json.Linq;  //Для работы с JSON
using System.Windows.Forms;  //Для использования MessageBox и PictureBox
using System.Net;  //Для скачивания изображения
using System.Drawing;  //Для использования графики и размеров
using System.Linq;  //Добавлено для методов расширения, таких как OrderByDescending

namespace NavisworksDefectPlugin7  //Добавляю пространство имен
{
    [Plugin("DefectMarkerPlugin", "DMP", ToolTip = "Marks defects using coordinates and images from specified folders", DisplayName = "Defect Marker")]
    public class DefectMarkerPlugin : AddInPlugin
    {
        //Папки для хранения изображений и JSON-файлов
        private readonly string IMAGES_FOLDER = "D:\\FlaskServer\\data\\images";
        private readonly string JSON_FOLDER = "D:\\FlaskServer\\data\\json";

        public override int Execute(params string[] parameters)
        {
            try
            {
                //Загружаю активный документ в Navisworks
                Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("Активный документ не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 1;
                }

                //Получаю последний JSON-файл из папки JSON_FOLDER
                string jsonFilePath = GetLatestJsonFile();
                if (string.IsNullOrEmpty(jsonFilePath) || !File.Exists(jsonFilePath))
                {
                    MessageBox.Show("Файл JSON не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 1;
                }

                //Парсю JSON-файл
                JObject json = JObject.Parse(File.ReadAllText(jsonFilePath));

                //Проверяю наличия координат в JSON
                if (json["coordinates"] == null || json["coordinates"]["x"] == null || json["coordinates"]["y"] == null || json["coordinates"]["z"] == null)
                {
                    MessageBox.Show("Координаты X, Y или Z не найдены в файле JSON.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 1;
                }

                //Получаею координаты X, Y, Z из JSON
                double x = (double)json["coordinates"]["x"];
                double y = (double)json["coordinates"]["y"];
                double z = (double)json["coordinates"]["z"];

                //Получаю путь к изображению дефекта из JSON
                string imageUrl = json["image_path"]?.ToString();
                if (string.IsNullOrEmpty(imageUrl))
                {
                    MessageBox.Show("Путь к изображению не найден в JSON.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 1;
                }

                //Проверяю наличие изображения в папке IMAGES_FOLDER
                string imagePath = GetImageFromFolder(imageUrl);
                if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath))
                {
                    MessageBox.Show("Файл изображения не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 1;
                }

                //Вывожу информацию для проверки
                MessageBox.Show($"Координаты дефекта: X = {x}, Y = {y}, Z = {z}\nИзображение найдено по адресу: {imagePath}");

                //Использую координат для поиска ближайшего элемента модели
                ModelItem nearestItem = FindNearestModelItem(doc, x, y, z);

                if (nearestItem != null)
                {
                    //Подсвечиваю найденный элемент
                    doc.CurrentSelection.Add(nearestItem);
                    MessageBox.Show($"Элемент найден: {nearestItem.DisplayName}");

                    //Отображаю изображение дефекта в отдельном окне
                    ShowDefectImage(imagePath);
                }
                else
                {
                    MessageBox.Show("Элемент не найден в модели.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }

            return 0;
        }

        //Метод для поиска последнего JSON-файла в папке JSON_FOLDER
        private string GetLatestJsonFile()
        {
            try
            {
                var directoryInfo = new DirectoryInfo(JSON_FOLDER);
                var latestFile = directoryInfo.GetFiles("*.json")
                                              .OrderByDescending(f => f.LastWriteTime)  //Используем метод OrderByDescending
                                              .FirstOrDefault();
                return latestFile?.FullName;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при поиске JSON-файлов: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        //Метод для поиска изображения в папке IMAGES_FOLDER
        private string GetImageFromFolder(string imageUrl)
        {
            try
            {
                string imageName = Path.GetFileName(imageUrl); //Извлекаем имя файла
                string imagePath = Path.Combine(IMAGES_FOLDER, imageName); //Путь к изображению
                return imagePath;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при поиске изображения: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        //Метод для поиска ближайшего элемента модели по координатам
        private ModelItem FindNearestModelItem(Document doc, double x, double y, double z)
        {
            ModelItem nearestItem = null;
            double minDistance = double.MaxValue;

            //Перебираю все элементы модели
            foreach (ModelItem item in doc.Models.First.RootItem.DescendantsAndSelf)
            {
                //Получаю BoundingBox для каждого элемента
                BoundingBox3D boundingBox = item.BoundingBox();

                if (boundingBox != null && !ArePointsEqual(boundingBox.Min, boundingBox.Max))
                {
                    //Вычисляю центр элемента
                    Point3D center = boundingBox.Center;

                    //Вычисляю расстояние до центра
                    double distance = Math.Sqrt(Math.Pow(center.X - x, 2) + Math.Pow(center.Y - y, 2) + Math.Pow(center.Z - z, 2));

                    //Нахожу ближайший элемент
                    if (distance < minDistance)
                    {
                        minDistance = distance;
                        nearestItem = item;
                    }
                }
            }

            return nearestItem;
        }

        //Метод для сравнения двух точек
        private bool ArePointsEqual(Point3D point1, Point3D point2)
        {
            return point1.X == point2.X && point1.Y == point2.Y && point1.Z == point2.Z;
        }

        //Метод для отображения изображения дефекта в отдельном окне
        private void ShowDefectImage(string imagePath)
        {
            Form imageForm = new Form();
            imageForm.Text = "Изображение дефекта";
            imageForm.Size = new Size(600, 600);  //Использую Size из System.Drawing

            PictureBox pictureBox = new PictureBox();
            pictureBox.ImageLocation = imagePath;
            pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox.Dock = DockStyle.Fill;

            imageForm.Controls.Add(pictureBox);
            imageForm.ShowDialog();
        }
    }
}
