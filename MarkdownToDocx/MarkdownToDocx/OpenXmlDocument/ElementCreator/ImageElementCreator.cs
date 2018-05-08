using System;
using DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;
using DP = DocumentFormat.OpenXml.Drawing.Pictures;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument.ElementCreator
{
    internal static class ImageElementCreator
    {
        public static Run CreateImageElement(string imageRelationshipId, long iamgeWidthInEmus, long imageHeightInEmus, string fileName, string description)
        {
            if (string.IsNullOrWhiteSpace(imageRelationshipId)) throw new ArgumentOutOfRangeException(nameof(imageRelationshipId), imageRelationshipId, "The image relationship ID is not valid.");
            if (iamgeWidthInEmus < 0) throw new ArgumentOutOfRangeException(nameof(iamgeWidthInEmus), iamgeWidthInEmus, "The image width is must be greater than equal 0.");
            if (imageHeightInEmus < 0) throw new ArgumentOutOfRangeException(nameof(imageHeightInEmus), imageHeightInEmus, "The image height is must be greater than equal 0.");
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentOutOfRangeException(nameof(fileName), fileName, "The image file name is not valid.");

            var imageId = DrawingElementIdGenerator.GetNewId();

            return new Run(
                new Drawing(
                    new DW.Inline(
                        new DW.Extent()
                        {
                            Cx = iamgeWidthInEmus,
                            Cy = imageHeightInEmus,
                        },
                        new DW.EffectExtent()
                        {
                            LeftEdge = 0,
                            TopEdge = 0,
                            RightEdge = 0,
                            BottomEdge = 0,
                        },
                        new DW.DocProperties()
                        {
                            Id = imageId,
                            Name = fileName,   // The name of this picture.
                            Description = string.IsNullOrEmpty(description) ? null : description,  // The alt text of this picture.
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new D.GraphicFrameLocks()
                            {
                                NoChangeAspect = true,
                            }
                        ),
                        new D.Graphic(
                            new D.GraphicData(
                                new DP.Picture(
                                    new DP.NonVisualPictureProperties(
                                        new DP.NonVisualDrawingProperties()
                                        {
                                            Id = imageId,
                                            Name = fileName,  // The file name of this picture.
                                        },
                                        new DP.NonVisualPictureDrawingProperties()
                                    ),
                                    new DP.BlipFill(
                                        new D.Blip()
                                        {
                                            Embed = imageRelationshipId,
                                        },
                                        new D.Stretch(
                                            new D.FillRectangle()
                                        )
                                    ),
                                    new DP.ShapeProperties(
                                        new D.Transform2D(
                                            new D.Offset()
                                            {
                                                X = 0,
                                                Y = 0,
                                            },
                                            new D.Extents()
                                            {
                                                Cx = iamgeWidthInEmus,
                                                Cy = imageHeightInEmus,
                                            }
                                        ),
                                        new D.PresetGeometry(
                                            new D.AdjustValueList()
                                        )
                                        {
                                            Preset = D.ShapeTypeValues.Rectangle,
                                        }
                                    )
                                )
                            )
                            {
                                Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
                            }
                        )
                    )
                    {
                        DistanceFromTop = 0U,
                        DistanceFromBottom = 0U,
                        DistanceFromLeft = 0U,
                        DistanceFromRight = 0U,
                    }
                )
            );
        }

        private static class DrawingElementIdGenerator
        {
            private static int NextDrawingElementId = 1;  // This ID start from 1.

            public static uint GetNewId()
            {
                //
                // The standard specifies that the ST_DrawingElementId type is an unsigned 32-bit integer.
                // But, Office treats the ST_DrawingElementId type as a signed 32-bit integer.
                // https://msdn.microsoft.com/en-us/library/ff533034(v=office.12).aspx
                //
                return (uint)NextDrawingElementId++;
            }
        }
    }
}
