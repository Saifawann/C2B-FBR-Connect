using C2B_FBR_Connect.Forms;
using C2B_FBR_Connect.Managers;
using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Services;
using C2B_FBR_Connect.Utilities;
using QBFC16Lib;
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace C2B_FBR_Connect.Utilities
{
    /// <summary>
    /// Utility class for image conversion and manipulation operations.
    /// Used for handling company logos and other images in the application.
    /// </summary>
    public static class ImageHelper
    {
        // Constants for default logo sizing
        private const int DEFAULT_MAX_WIDTH = 500;
        private const int DEFAULT_MAX_HEIGHT = 500;
        private const long DEFAULT_MAX_FILE_SIZE = 2 * 1024 * 1024; // 2MB

        /// <summary>
        /// Convert Image to byte array in PNG format
        /// </summary>
        /// <param name="image">The image to convert</param>
        /// <returns>Byte array representation of the image, or null if input is null</returns>
        public static byte[] ImageToByteArray(Image image)
        {
            if (image == null) return null;

            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    // Save as PNG for lossless quality
                    image.Save(ms, ImageFormat.Png);
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error converting image to byte array: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Convert byte array to Image object
        /// </summary>
        /// <param name="byteArray">Byte array containing image data</param>
        /// <returns>Image object, or null if input is null or invalid</returns>
        public static Image ByteArrayToImage(byte[] byteArray)
        {
            if (byteArray == null || byteArray.Length == 0) return null;

            try
            {
                using (MemoryStream ms = new MemoryStream(byteArray))
                {
                    // Create a copy to avoid issues with stream disposal
                    return new Bitmap(Image.FromStream(ms));
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error converting byte array to image: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Resize image to fit within maximum dimensions while maintaining aspect ratio
        /// </summary>
        /// <param name="image">Original image to resize</param>
        /// <param name="maxWidth">Maximum width in pixels</param>
        /// <param name="maxHeight">Maximum height in pixels</param>
        /// <returns>Resized image, or original if already within bounds</returns>
        public static Image ResizeImage(Image image, int maxWidth, int maxHeight)
        {
            if (image == null)
                throw new ArgumentNullException(nameof(image));

            if (maxWidth <= 0 || maxHeight <= 0)
                throw new ArgumentException("Maximum dimensions must be positive values");

            // Calculate scaling ratio while maintaining aspect ratio
            double ratioX = (double)maxWidth / image.Width;
            double ratioY = (double)maxHeight / image.Height;
            double ratio = Math.Min(ratioX, ratioY);

            // Don't upscale if image is already smaller than max dimensions
            if (ratio >= 1.0)
            {
                return new Bitmap(image);
            }

            int newWidth = (int)(image.Width * ratio);
            int newHeight = (int)(image.Height * ratio);

            try
            {
                Bitmap resizedImage = new Bitmap(newWidth, newHeight);

                using (Graphics graphics = Graphics.FromImage(resizedImage))
                {
                    // Set high-quality rendering options
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    graphics.CompositingQuality = CompositingQuality.HighQuality;

                    // Draw the resized image
                    graphics.DrawImage(image, 0, 0, newWidth, newHeight);
                }

                return resizedImage;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error resizing image: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Load image from file, resize it, and convert to byte array
        /// </summary>
        /// <param name="filePath">Full path to the image file</param>
        /// <param name="maxWidth">Maximum width (default: 500px)</param>
        /// <param name="maxHeight">Maximum height (default: 500px)</param>
        /// <returns>Byte array of resized image, or null if file doesn't exist</returns>
        public static byte[] LoadAndResizeImage(string filePath, int maxWidth = DEFAULT_MAX_WIDTH, int maxHeight = DEFAULT_MAX_HEIGHT)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path cannot be null or empty", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Image file not found", filePath);

            // Check file size
            FileInfo fileInfo = new FileInfo(filePath);
            if (fileInfo.Length > DEFAULT_MAX_FILE_SIZE)
            {
                throw new Exception($"Image file size ({fileInfo.Length / 1024 / 1024}MB) exceeds maximum allowed size ({DEFAULT_MAX_FILE_SIZE / 1024 / 1024}MB)");
            }

            try
            {
                using (Image originalImage = Image.FromFile(filePath))
                using (Image resizedImage = ResizeImage(originalImage, maxWidth, maxHeight))
                {
                    return ImageToByteArray(resizedImage);
                }
            }
            catch (OutOfMemoryException)
            {
                throw new Exception("Invalid image file or unsupported format");
            }
            catch (Exception ex)
            {
                throw new Exception($"Error loading and resizing image: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Validate if a file is a valid image format
        /// </summary>
        /// <param name="filePath">Path to the file</param>
        /// <returns>True if valid image, false otherwise</returns>
        public static bool IsValidImageFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return false;

            try
            {
                using (Image img = Image.FromFile(filePath))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Get image dimensions without loading the full image
        /// </summary>
        /// <param name="filePath">Path to the image file</param>
        /// <returns>Tuple with width and height, or (0,0) if invalid</returns>
        public static (int width, int height) GetImageDimensions(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return (0, 0);

            try
            {
                using (Image img = Image.FromFile(filePath))
                {
                    return (img.Width, img.Height);
                }
            }
            catch
            {
                return (0, 0);
            }
        }

        /// <summary>
        /// Get human-readable file size
        /// </summary>
        /// <param name="filePath">Path to the file</param>
        /// <returns>Formatted file size string (e.g., "1.5 MB")</returns>
        public static string GetFileSizeFormatted(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return "0 B";

            FileInfo fileInfo = new FileInfo(filePath);
            long bytes = fileInfo.Length;

            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;

            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }

            return $"{len:0.##} {sizes[order]}";
        }

        /// <summary>
        /// Create a placeholder image for missing logos
        /// </summary>
        /// <param name="width">Width of placeholder</param>
        /// <param name="height">Height of placeholder</param>
        /// <param name="text">Text to display (optional)</param>
        /// <returns>Byte array of placeholder image</returns>
        public static byte[] CreatePlaceholderImage(int width = 200, int height = 100, string text = "No Logo")
        {
            try
            {
                Bitmap placeholder = new Bitmap(width, height);

                using (Graphics graphics = Graphics.FromImage(placeholder))
                {
                    // Fill background
                    graphics.Clear(Color.FromArgb(240, 240, 240));

                    // Draw border
                    using (Pen borderPen = new Pen(Color.FromArgb(200, 200, 200), 2))
                    {
                        graphics.DrawRectangle(borderPen, 1, 1, width - 2, height - 2);
                    }

                    // Draw text
                    if (!string.IsNullOrEmpty(text))
                    {
                        using (Font font = new Font("Arial", 12, FontStyle.Regular))
                        using (SolidBrush brush = new SolidBrush(Color.FromArgb(150, 150, 150)))
                        {
                            StringFormat format = new StringFormat
                            {
                                Alignment = StringAlignment.Center,
                                LineAlignment = StringAlignment.Center
                            };

                            graphics.DrawString(text, font, brush, new RectangleF(0, 0, width, height), format);
                        }
                    }
                }

                return ImageToByteArray(placeholder);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error creating placeholder image: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Compress image by reducing quality (JPEG format)
        /// </summary>
        /// <param name="image">Image to compress</param>
        /// <param name="quality">Quality level (0-100, default 85)</param>
        /// <returns>Compressed image as byte array</returns>
        public static byte[] CompressImage(Image image, long quality = 85L)
        {
            if (image == null)
                throw new ArgumentNullException(nameof(image));

            if (quality < 0 || quality > 100)
                throw new ArgumentOutOfRangeException(nameof(quality), "Quality must be between 0 and 100");

            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    // Get JPEG codec
                    ImageCodecInfo jpegCodec = GetEncoderInfo("image/jpeg");

                    // Set quality parameter
                    EncoderParameters encoderParameters = new EncoderParameters(1);
                    encoderParameters.Param[0] = new EncoderParameter(Encoder.Quality, quality);

                    // Save with compression
                    image.Save(ms, jpegCodec, encoderParameters);
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error compressing image: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Get image encoder for a specific MIME type
        /// </summary>
        private static ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.MimeType == mimeType)
                    return codec;
            }
            return null;
        }
    }
}