import { useState, useRef, useEffect } from 'react';
import { Stage, Layer, Image as KonvaImage, Rect, Text, Group } from 'react-konva';
import type { KonvaEventObject } from 'konva/lib/Node';
import { saveAs } from 'file-saver';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  HeadingLevel,
  BorderStyle,
  VerticalAlign,
} from 'docx';
import JSZip from 'jszip';

interface Annotation {
  x: number;
  y: number;
  width: number;
  height: number;
  defectType: string;
  notes: string;
  isLocationPlan?: boolean;
  isLocationExtent?: boolean;
}

interface ImageData {
  file: File;
  img: HTMLImageElement;
  annotations: Annotation[];
  type: 'visual' | 'infrared' | 'hyperspectral';
  locationPlan?: File;
  locationPlanImg?: HTMLImageElement;
  locationPlanAnnotations?: Annotation[];
}

interface FormData {
  location: string;
  locationMap: File | null;
  drawingFile: File | null;
}

const defectTypes = [
  'Crack',
  'Spalling',
  'Seepage',
  'Stain Tiles',
  'Chipping Tiles',
  'Rust',
  'Delamination',
  'Aged Sealant',
  'Hot Abnormal',
  'Cold Abnormal',
  'Dirt'
];

const getDefectColor = (defectType: string, alpha: number = 1): string => {
  const colors: { [key: string]: string } = {
    'Crack': `rgba(255, 0, 0, ${alpha})`,          // Red
    'Spalling': `rgba(255, 165, 0, ${alpha})`,     // Orange
    'Seepage': `rgba(0, 0, 255, ${alpha})`,        // Blue
    'Stain Tiles': `rgba(128, 0, 128, ${alpha})`,  // Purple
    'Chipping Tiles': `rgba(255, 192, 203, ${alpha})`, // Pink
    'Rust': `rgba(139, 69, 19, ${alpha})`,         // Brown
    'Delamination': `rgba(0, 128, 0, ${alpha})`,   // Green
    'Aged Sealant': `rgba(255, 255, 0, ${alpha})`, // Yellow
    'Hot Abnormal': `rgba(255, 0, 255, ${alpha})`, // Magenta
    'Cold Abnormal': `rgba(0, 255, 255, ${alpha})`, // Cyan
    'Dirt': `rgba(210, 180, 140, ${alpha})`        // Tan/Beige
  };
  return colors[defectType] || `rgba(128, 128, 128, ${alpha})`; // Default gray
};

const Canvas = ({
  image,
  annotations,
  setAnnotations,
  onSelectAnnotation,
  onDimensionsChange,
  isParallel = false,
  onLocationPlanUpload,
  isLocationPlan = false
}: {
  image: HTMLImageElement | null;
  annotations: Annotation[];
  setAnnotations: (annotations: Annotation[]) => void;
  onSelectAnnotation: (index: number) => void;
  onDimensionsChange: (dimensions: { width: number; height: number }) => void;
  isParallel?: boolean;
  onLocationPlanUpload?: (file: File) => void;
  isLocationPlan?: boolean;
}) => {
  const [drawing, setDrawing] = useState(false);
  const [startPoint, setStartPoint] = useState<{ x: number; y: number } | null>(null);
  const [selectedDefectType, setSelectedDefectType] = useState<string>('');
  const [drawingMode, setDrawingMode] = useState<'defect' | 'locationExtent'>('defect');
  const stageRef = useRef<any>(null);
  const [dimensions, setDimensions] = useState({ width: 600, height: 400 });
  const [imageLoaded, setImageLoaded] = useState(false);
  const [currentRect, setCurrentRect] = useState<{ x: number; y: number; width: number; height: number } | null>(null);
  const [scale, setScale] = useState(1);
  const [position, setPosition] = useState({ x: 0, y: 0 });
  const [isDragging, setIsDragging] = useState(false);
  const lastPos = useRef<{ x: number; y: number }>({ x: 0, y: 0 });
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (image) {
      const updateDimensions = () => {
        const maxWidth = Math.min(800, window.innerWidth - 40);
        const maxHeight = Math.min(600, window.innerHeight - 200);
        
        let width = image.width;
        let height = image.height;
        
        if (width > maxWidth) {
          const scale = maxWidth / width;
          width = maxWidth;
          height = height * scale;
        }
        
        if (height > maxHeight) {
          const scale = maxHeight / height;
          height = maxHeight;
          width = width * scale;
        }
        
        setDimensions({ width, height });
        setImageLoaded(true);
        onDimensionsChange({ width, height });
      };

      if (image.complete) {
        updateDimensions();
      } else {
        image.onload = updateDimensions;
      }
    }
  }, [image, onDimensionsChange]);

  const handleMouseDown = (e: KonvaEventObject<MouseEvent>) => {
    if (!drawing) {
      setIsDragging(true);
      return;
    }
    const stage = e.target.getStage();
    if (stage) {
      const pos = stage.getPointerPosition();
      if (pos) {
        setStartPoint(pos);
        setCurrentRect({
          x: pos.x,
          y: pos.y,
          width: 0,
          height: 0
        });
      }
    }
  };

  const handleMouseMove = (e: KonvaEventObject<MouseEvent>) => {
    if (isDragging && !drawing) {
      const stage = e.target.getStage();
      if (stage) {
        const currentPos = stage.getPointerPosition();
        if (currentPos) {
          setPosition({
            x: position.x + (currentPos.x - lastPos.current.x),
            y: position.y + (currentPos.y - lastPos.current.y)
          });
          lastPos.current = currentPos;
        }
      }
      return;
    }
    if (!drawing || !startPoint || !currentRect) return;
    const stage = e.target.getStage();
    if (stage) {
      const pos = stage.getPointerPosition();
      if (pos) {
        setCurrentRect({
          x: Math.min(pos.x, startPoint.x),
          y: Math.min(pos.y, startPoint.y),
          width: Math.abs(pos.x - startPoint.x),
          height: Math.abs(pos.y - startPoint.y)
        });
      }
    }
  };

  const handleMouseUp = () => {
    setIsDragging(false);
    if (drawing && startPoint && currentRect && (currentRect.width > 5 || currentRect.height > 5)) {
      const newAnnotation = {
        ...currentRect,
        defectType: drawingMode === 'defect' ? selectedDefectType : 'Location Extent',
        notes: '',
        isLocationExtent: drawingMode === 'locationExtent'
      };
      setAnnotations([...annotations, newAnnotation]);
      setStartPoint(null);
      setCurrentRect(null);
      setDrawing(false);
      setSelectedDefectType('');
      setDrawingMode('defect');
    }
  };

  const handleWheel = (e: KonvaEventObject<WheelEvent>) => {
    if (isParallel) return;
    e.evt.preventDefault();
    const stage = e.target.getStage();
    if (!stage) return;

    const oldScale = scale;
    const pointer = stage.getPointerPosition();
    if (!pointer) return;

    const mousePointTo = {
      x: (pointer.x - position.x) / oldScale,
      y: (pointer.y - position.y) / oldScale,
    };

    const newScale = e.evt.deltaY < 0 ? oldScale * 1.1 : oldScale / 1.1;
    setScale(newScale);

    const newPos = {
      x: pointer.x - mousePointTo.x * newScale,
      y: pointer.y - mousePointTo.y * newScale,
    };
    setPosition(newPos);
  };

  const resetZoom = () => {
    setScale(1);
    setPosition({ x: 0, y: 0 });
  };

  const startDrawing = (defectType: string) => {
    setSelectedDefectType(defectType);
    setDrawing(true);
  };

  const handleLocationPlanUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file && onLocationPlanUpload) {
      onLocationPlanUpload(file);
    }
  };

  const startLocationExtentDrawing = () => {
    setDrawingMode('locationExtent');
    setDrawing(true);
  };

  return (
    <div className="relative border rounded overflow-hidden bg-gray-100">
      <div className="absolute top-2 left-2 z-10 flex flex-wrap gap-2 p-2 bg-white/80 rounded shadow">
        {drawing ? (
          <button
            className="bg-red-500 text-white px-4 py-2 rounded hover:bg-red-600"
            onClick={() => {
              setDrawing(false);
              setSelectedDefectType('');
              setStartPoint(null);
              setCurrentRect(null);
              setDrawingMode('defect');
            }}
          >
            Cancel Drawing
          </button>
        ) : !isParallel ? (
          <div className="space-y-2">
            {!isLocationPlan ? (
              <>
                <div className="grid grid-cols-2 gap-2">
                  {defectTypes.map((type) => (
                    <button
                      key={type}
                      className={`px-3 py-1.5 rounded text-sm font-medium
                        ${selectedDefectType === type 
                          ? 'bg-blue-600 text-white' 
                          : 'bg-blue-100 text-blue-700 hover:bg-blue-200'}`}
                      onClick={() => {
                        setDrawingMode('defect');
                        startDrawing(type);
                      }}
                    >
                      {type}
                    </button>
                  ))}
                </div>
                <div className="flex gap-2">
                  <button
                    className="px-3 py-1.5 rounded text-sm font-medium bg-purple-100 text-purple-700 hover:bg-purple-200"
                    onClick={() => fileInputRef.current?.click()}
                  >
                    Insert Location Plan
                  </button>
                  <input
                    type="file"
                    ref={fileInputRef}
                    accept="image/*"
                    onChange={handleLocationPlanUpload}
                    className="hidden"
                  />
                </div>
              </>
            ) : (
              <button
                className="px-3 py-1.5 rounded text-sm font-medium bg-green-100 text-green-700 hover:bg-green-200"
                onClick={startLocationExtentDrawing}
              >
                Label Location Extent
              </button>
            )}
          </div>
        ) : null}
      </div>
      
      {!isParallel && (
        <div className="absolute top-2 right-2 z-10 flex gap-2">
          <button
            className="bg-white/80 px-3 py-1.5 rounded text-sm font-medium hover:bg-white"
            onClick={() => setScale(scale * 1.1)}
          >
            +
          </button>
          <button
            className="bg-white/80 px-3 py-1.5 rounded text-sm font-medium hover:bg-white"
            onClick={() => setScale(scale / 1.1)}
          >
            -
          </button>
          <button
            className="bg-white/80 px-3 py-1.5 rounded text-sm font-medium hover:bg-white"
            onClick={resetZoom}
          >
            Reset
          </button>
        </div>
      )}
      
      {image && imageLoaded ? (
        <div className="relative">
          <Stage
            width={dimensions.width}
            height={dimensions.height}
            ref={stageRef}
            onMouseDown={!isParallel ? handleMouseDown : undefined}
            onMouseMove={!isParallel ? handleMouseMove : undefined}
            onMouseUp={!isParallel ? handleMouseUp : undefined}
            onWheel={!isParallel ? handleWheel : undefined}
            scaleX={scale}
            scaleY={scale}
            x={position.x}
            y={position.y}
            draggable={!drawing && !isParallel}
          >
            <Layer>
              <KonvaImage
                image={image}
                width={dimensions.width}
                height={dimensions.height}
              />
              {annotations.map((ann, index) => (
                <Group key={index} onClick={() => !isParallel && onSelectAnnotation(index)}>
                  <Rect
                    x={ann.x}
                    y={ann.y}
                    width={ann.width}
                    height={ann.height}
                    stroke={ann.isLocationExtent ? 'rgba(0, 255, 0, 0.8)' : getDefectColor(ann.defectType)}
                    strokeWidth={ann.isLocationExtent ? 3 : 2}
                    fill="transparent"
                    dash={ann.isLocationExtent ? [10, 5] : []}
                  />
                  <Rect
                    x={ann.x}
                    y={ann.y - 25}
                    width={ann.defectType.length * 10 + 10}
                    height={25}
                    fill="white"
                    opacity={0.8}
                  />
                  <Text
                    x={ann.x + 5}
                    y={ann.y - 20}
                    text={ann.defectType}
                    fontSize={16}
                    fill="black"
                    padding={2}
                  />
                </Group>
              ))}
              {currentRect && !isParallel && (
                <Rect
                  x={currentRect.x}
                  y={currentRect.y}
                  width={currentRect.width}
                  height={currentRect.height}
                  stroke={drawingMode === 'locationExtent' ? 'rgba(0, 255, 0, 0.8)' : getDefectColor(selectedDefectType)}
                  strokeWidth={drawingMode === 'locationExtent' ? 3 : 2}
                  fill="transparent"
                  dash={drawingMode === 'locationExtent' ? [10, 5] : []}
                />
              )}
            </Layer>
          </Stage>
        </div>
      ) : (
        <div className="flex items-center justify-center" style={{ width: 600, height: 400 }}>
          <p className="text-gray-500">{image ? 'Loading image...' : 'No image selected'}</p>
        </div>
      )}
    </div>
  );
};

function App() {
  const [images, setImages] = useState<ImageData[]>([]);
  const [currentImageIndex, setCurrentImageIndex] = useState(0);
  const [selectedAnnotationIndex, setSelectedAnnotationIndex] = useState<number | null>(null);
  const [formData, setFormData] = useState<FormData>({
    location: '',
    locationMap: null,
    drawingFile: null,
  });
  const [canvasDimensions, setCanvasDimensions] = useState({ width: 600, height: 400 });

  // Add new state for parallel images
  const [parallelImages, setParallelImages] = useState<{
    infrared: ImageData[];
    hyperspectral: ImageData[];
  }>({
    infrared: [],
    hyperspectral: []
  });

  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>, type: 'visual' | 'infrared' | 'hyperspectral' = 'visual') => {
    const files = e.target.files;
    if (!files) return;

    const newImages: ImageData[] = [];
    Array.from(files).forEach(file => {
      const img = new Image();
      img.src = URL.createObjectURL(file);
      newImages.push({
        file,
        img,
        annotations: [],
        type
      });
    });

    if (type === 'visual') {
      setImages([...images, ...newImages]);
    } else {
      setParallelImages(prev => ({
        ...prev,
        [type]: [...prev[type], ...newImages]
      }));
    }
  };

  const handleFileUpload = (type: 'locationMap' | 'drawingFile') => (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      setFormData(prev => ({
        ...prev,
        [type]: file
      }));
    }
  };

  const handleAnnotationUpdate = (_index: number, updates: Partial<Annotation>) => {
    const newImages = [...images];
    if (selectedAnnotationIndex !== null && newImages[currentImageIndex]) {
      const currentAnnotations = newImages[currentImageIndex].annotations;
      currentAnnotations[selectedAnnotationIndex] = {
        ...currentAnnotations[selectedAnnotationIndex],
        ...updates
      };
      setImages(newImages);
    }
  };

  const saveAnnotatedImage = async (imageData: ImageData, _index: number): Promise<Blob | null> => {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    if (!ctx) return null;

    // Set canvas size to match image
    canvas.width = imageData.img.width;
    canvas.height = imageData.img.height;

    // Draw the original image
    ctx.drawImage(imageData.img, 0, 0);

    // Draw annotations
    imageData.annotations.forEach(ann => {
      const scaleX = imageData.img.width / canvasDimensions.width;
      const scaleY = imageData.img.height / canvasDimensions.height;
      
      ctx.strokeStyle = getDefectColor(ann.defectType);
      ctx.fillStyle = getDefectColor(ann.defectType, 0.3);
      
      ctx.strokeRect(
        ann.x * scaleX,
        ann.y * scaleY,
        ann.width * scaleX,
        ann.height * scaleY
      );
      ctx.fillRect(
        ann.x * scaleX,
        ann.y * scaleY,
        ann.width * scaleX,
        ann.height * scaleY
      );

      // Add label
      ctx.font = '16px Arial';
      ctx.fillStyle = 'white';
      ctx.fillRect(
        ann.x * scaleX,
        (ann.y * scaleY) - 25,
        ctx.measureText(ann.defectType).width + 10,
        25
      );
      ctx.fillStyle = 'black';
      ctx.fillText(ann.defectType, ann.x * scaleX + 5, (ann.y * scaleY) - 5);
    });

    // Return the blob
    return new Promise((resolve) => {
      canvas.toBlob((blob) => {
        resolve(blob);
      }, 'image/png');
    });
  };

  const saveCurrentImage = async () => {
    if (!images[currentImageIndex]) return;

    const currentImage = images[currentImageIndex];
    const blob = await saveAnnotatedImage(currentImage, currentImageIndex);
    if (blob) {
      const fileName = `${currentImage.file.name.replace(/\.[^/.]+$/, '')}_annotated_${currentImageIndex + 1}.png`;
      saveAs(blob, fileName);
    }
  };

  const saveAllImages = async () => {
    if (images.length === 0) return;

    const zip = new JSZip();
    const annotatedFolder = zip.folder("annotated_images");
    
    if (!annotatedFolder) return;

    // Show loading state
    const button = document.getElementById('save-all-button') as HTMLButtonElement;
    if (button) {
      button.disabled = true;
      button.innerHTML = `
        <svg class="animate-spin h-5 w-5 mr-2" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
          <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
          <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
        </svg>
        Processing...
      `;
    }

    try {
      // Process all images
      for (let i = 0; i < images.length; i++) {
        const blob = await saveAnnotatedImage(images[i], i);
        if (blob) {
          const fileName = `${images[i].file.name.replace(/\.[^/.]+$/, '')}_annotated_${i + 1}.png`;
          annotatedFolder.file(fileName, blob);
        }
      }

      // Generate and save the zip file
      const content = await zip.generateAsync({ type: "blob" });
      const date = new Date().toISOString().split('T')[0];
      saveAs(content, `annotated_images_${date}.zip`);
    } catch (error) {
      console.error('Error creating zip file:', error);
    } finally {
      // Reset button state
      if (button) {
        button.disabled = false;
        button.innerHTML = `
          <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5 mr-2 text-gray-500" viewBox="0 0 20 20" fill="currentColor">
            <path fill-rule="evenodd" d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zm3.293-7.707a1 1 0 011.414 0L9 10.586V3a1 1 0 112 0v7.586l1.293-1.293a1 1 0 111.414 1.414l-3 3a1 1 0 01-1.414 0l-3-3a1 1 0 010-1.414z" clip-rule="evenodd" />
          </svg>
          Save All Images
        `;
      }
    }
  };

  const goToNextImage = () => {
    if (currentImageIndex < images.length - 1) {
      setCurrentImageIndex(prev => prev + 1);
      setSelectedAnnotationIndex(null);
    }
  };

  const goToPreviousImage = () => {
    if (currentImageIndex > 0) {
      setCurrentImageIndex(prev => prev - 1);
      setSelectedAnnotationIndex(null);
    }
  };

  const currentImage = images[currentImageIndex]?.img || null;
  const currentAnnotations = images[currentImageIndex]?.annotations || [];

  const generateReport = async () => {
    if (images.length === 0) return;

    // Create a list of defects with IDs
    const defectsList: { 
      type: string; 
      id: string;
      notes: string; 
      imageIndex: number;
    }[] = [];

    // Process all images to generate IDs
    images.forEach((image, imageIndex) => {
      image.annotations.forEach(ann => {
        // Generate ID based on position and image index
        const id = `${ann.defectType.replace(/\s+/g, '')}${(imageIndex + 1).toString().padStart(2, '0')}${ann.x.toString().padStart(3, '0')}${ann.y.toString().padStart(3, '0')}`;
        defectsList.push({
          type: ann.defectType,
          id,
          notes: ann.notes,
          imageIndex,
        });
      });
    });
    
    // Create sections for each image with its annotations
    const sections: any[] = [];
    
    // Add report metadata section
    sections.push(
      new Paragraph({
        text: "Inspection Report",
        heading: HeadingLevel.HEADING_1,
      }),
      new Paragraph({
        text: "",
      }),
      new Table({
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [new Paragraph("Location:")],
                width: {
                  size: 20,
                  type: WidthType.PERCENTAGE,
                },
              }),
              new TableCell({
                children: [new Paragraph(formData.location)],
              }),
            ],
          }),
        ],
      }),
      new Paragraph({ text: "" }),
      new Paragraph({
        text: "Location Details",
        heading: HeadingLevel.HEADING_2,
      }),
      new Table({
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
          left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
          right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
          insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [new Paragraph("Location:")],
                width: {
                  size: 20,
                  type: WidthType.PERCENTAGE,
                },
              }),
              new TableCell({
                children: [new Paragraph(formData.location)],
              }),
            ],
          }),
        ],
      }),
    );

    // Add location map and drawing in a table
    const locationImagesTable = new Table({
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  text: "Location Map",
                  alignment: AlignmentType.CENTER,
                }),
              ],
            }),
            new TableCell({
              children: [
                new Paragraph({
                  text: "Drawing",
                  alignment: AlignmentType.CENTER,
                }),
              ],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  children: formData.locationMap ? [
                    new ImageRun({
                      data: await processImageToBase64(formData.locationMap),
                      transformation: {
                        width: 250,
                        height: 200,
                      },
                      type: 'png',
                    }),
                  ] : [new TextRun("No location map provided")],
                  alignment: AlignmentType.CENTER,
                }),
              ],
            }),
            new TableCell({
              children: [
                new Paragraph({
                  children: formData.drawingFile ? [
                    new ImageRun({
                      data: await processImageToBase64(formData.drawingFile),
                      transformation: {
                        width: 250,
                        height: 200,
                      },
                      type: 'png',
                    }),
                  ] : [new TextRun("No drawing provided")],
                  alignment: AlignmentType.CENTER,
                }),
              ],
              verticalAlign: VerticalAlign.CENTER,
            }),
          ],
        }),
      ],
    });

    sections.push(locationImagesTable, new Paragraph({ text: "" }));

    // Process each image
    for (let i = 0; i < images.length; i++) {
      const currentImage = images[i];
      const imageDefects = defectsList.filter(d => d.imageIndex === i);
      
      // Create canvas with annotations for visual image
      const visualCanvas = document.createElement('canvas');
      const visualCtx = visualCanvas.getContext('2d');
      if (!visualCtx) continue;

      visualCanvas.width = currentImage.img.width;
      visualCanvas.height = currentImage.img.height;

      // Draw visual image and annotations
      visualCtx.drawImage(currentImage.img, 0, 0);
      currentImage.annotations.forEach((ann, index) => {
        const scaleX = currentImage.img.width / canvasDimensions.width;
        const scaleY = currentImage.img.height / canvasDimensions.height;
        
        // Draw rectangle
        visualCtx.strokeStyle = getDefectColor(ann.defectType);
        visualCtx.strokeRect(
          ann.x * scaleX,
          ann.y * scaleY,
          ann.width * scaleX,
          ann.height * scaleY
        );

        // Add white background for label
        visualCtx.font = '16px Arial';
        const textWidth = visualCtx.measureText(ann.defectType).width;
        visualCtx.fillStyle = 'rgba(255, 255, 255, 0.8)';
        visualCtx.fillRect(
          ann.x * scaleX,
          (ann.y * scaleY) - 25,
          textWidth + 10,
          25
        );

        // Add label text
        visualCtx.fillStyle = 'black';
        visualCtx.fillText(ann.defectType, ann.x * scaleX + 5, (ann.y * scaleY) - 5);
      });

      // Create canvas with annotations for location plan
      const locationPlanCanvas = document.createElement('canvas');
      const locationPlanCtx = locationPlanCanvas.getContext('2d');
      if (locationPlanCtx && currentImage.locationPlanImg) {
        locationPlanCanvas.width = currentImage.locationPlanImg.width;
        locationPlanCanvas.height = currentImage.locationPlanImg.height;

        // Draw location plan image
        locationPlanCtx.drawImage(currentImage.locationPlanImg, 0, 0);

        // Draw location extent annotations
        currentImage.locationPlanAnnotations?.forEach((ann) => {
          const scaleX = currentImage.locationPlanImg!.width / canvasDimensions.width;
          const scaleY = currentImage.locationPlanImg!.height / canvasDimensions.height;
          
          // Draw rectangle
          locationPlanCtx.strokeStyle = 'rgba(0, 255, 0, 0.8)';
          locationPlanCtx.setLineDash([10, 5]);
          locationPlanCtx.lineWidth = 3;
          locationPlanCtx.strokeRect(
            ann.x * scaleX,
            ann.y * scaleY,
            ann.width * scaleX,
            ann.height * scaleY
          );

          // Add white background for label
          locationPlanCtx.font = '16px Arial';
          const textWidth = locationPlanCtx.measureText('Location Extent').width;
          locationPlanCtx.fillStyle = 'rgba(255, 255, 255, 0.8)';
          locationPlanCtx.fillRect(
            ann.x * scaleX,
            (ann.y * scaleY) - 25,
            textWidth + 10,
            25
          );

          // Add label text
          locationPlanCtx.fillStyle = 'black';
          locationPlanCtx.fillText('Location Extent', ann.x * scaleX + 5, (ann.y * scaleY) - 5);
        });
      }

      // Create canvas with annotations for infrared image
      const infraredCanvas = document.createElement('canvas');
      const infraredCtx = infraredCanvas.getContext('2d');
      if (infraredCtx && parallelImages.infrared[i]) {
        infraredCanvas.width = parallelImages.infrared[i].img.width;
        infraredCanvas.height = parallelImages.infrared[i].img.height;

        // Draw infrared image and annotations
        infraredCtx.drawImage(parallelImages.infrared[i].img, 0, 0);
        currentImage.annotations.forEach((ann, index) => {
          const scaleX = parallelImages.infrared[i].img.width / canvasDimensions.width;
          const scaleY = parallelImages.infrared[i].img.height / canvasDimensions.height;
          
          // Draw rectangle
          infraredCtx.strokeStyle = getDefectColor(ann.defectType);
          infraredCtx.strokeRect(
            ann.x * scaleX,
            ann.y * scaleY,
            ann.width * scaleX,
            ann.height * scaleY
          );

          // Add white background for label
          infraredCtx.font = '16px Arial';
          const textWidth = infraredCtx.measureText(ann.defectType).width;
          infraredCtx.fillStyle = 'rgba(255, 255, 255, 0.8)';
          infraredCtx.fillRect(
            ann.x * scaleX,
            (ann.y * scaleY) - 25,
            textWidth + 10,
            25
          );

          // Add label text
          infraredCtx.fillStyle = 'black';
          infraredCtx.fillText(ann.defectType, ann.x * scaleX + 5, (ann.y * scaleY) - 5);
        });
      }

      // Create canvas with annotations for hyperspectral image
      const hyperspectralCanvas = document.createElement('canvas');
      const hyperspectralCtx = hyperspectralCanvas.getContext('2d');
      if (hyperspectralCtx && parallelImages.hyperspectral[i]) {
        hyperspectralCanvas.width = parallelImages.hyperspectral[i].img.width;
        hyperspectralCanvas.height = parallelImages.hyperspectral[i].img.height;

        // Draw hyperspectral image and annotations
        hyperspectralCtx.drawImage(parallelImages.hyperspectral[i].img, 0, 0);
        currentImage.annotations.forEach((ann, index) => {
          const scaleX = parallelImages.hyperspectral[i].img.width / canvasDimensions.width;
          const scaleY = parallelImages.hyperspectral[i].img.height / canvasDimensions.height;
          
          // Draw rectangle
          hyperspectralCtx.strokeStyle = getDefectColor(ann.defectType);
          hyperspectralCtx.strokeRect(
            ann.x * scaleX,
            ann.y * scaleY,
            ann.width * scaleX,
            ann.height * scaleY
          );

          // Add white background for label
          hyperspectralCtx.font = '16px Arial';
          const textWidth = hyperspectralCtx.measureText(ann.defectType).width;
          hyperspectralCtx.fillStyle = 'rgba(255, 255, 255, 0.8)';
          hyperspectralCtx.fillRect(
            ann.x * scaleX,
            (ann.y * scaleY) - 25,
            textWidth + 10,
            25
          );

          // Add label text
          hyperspectralCtx.fillStyle = 'black';
          hyperspectralCtx.fillText(ann.defectType, ann.x * scaleX + 5, (ann.y * scaleY) - 5);
        });
      }

      // Add image section with table
      sections.push(
        new Paragraph({
          text: `Image ${i + 1}: ${currentImage.file.name}`,
          heading: HeadingLevel.HEADING_2,
        }),
        new Table({
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  children: [new Paragraph({ text: "Visual Image" })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new ImageRun({
                          data: Uint8Array.from(atob(visualCanvas.toDataURL('image/png').split(',')[1]), c => c.charCodeAt(0)),
                          transformation: {
                            width: 600,
                            height: (600 * visualCanvas.height) / visualCanvas.width,
                          },
                          type: 'png'
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [new Paragraph({ text: "Location Plan" })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: currentImage.locationPlanImg ? [
                        new ImageRun({
                          data: Uint8Array.from(atob(locationPlanCanvas.toDataURL('image/png').split(',')[1]), c => c.charCodeAt(0)),
                          transformation: {
                            width: 600,
                            height: (600 * locationPlanCanvas.height) / locationPlanCanvas.width,
                          },
                          type: 'png'
                        }),
                      ] : [new TextRun("No location plan available")],
                    }),
                  ],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [new Paragraph({ text: "Infrared Image" })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: parallelImages.infrared.length > 0 ? [
                        new ImageRun({
                          data: Uint8Array.from(atob(infraredCanvas.toDataURL('image/png').split(',')[1]), c => c.charCodeAt(0)),
                          transformation: {
                            width: 600,
                            height: (600 * infraredCanvas.height) / infraredCanvas.width,
                          },
                          type: 'png'
                        }),
                      ] : [new TextRun("No infrared image available")],
                    }),
                  ],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [new Paragraph({ text: "Hyperspectral Image" })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: parallelImages.hyperspectral.length > 0 ? [
                        new ImageRun({
                          data: Uint8Array.from(atob(hyperspectralCanvas.toDataURL('image/png').split(',')[1]), c => c.charCodeAt(0)),
                          transformation: {
                            width: 600,
                            height: (600 * hyperspectralCanvas.height) / hyperspectralCanvas.width,
                          },
                          type: 'png'
                        }),
                      ] : [new TextRun("No hyperspectral image available")],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
        new Table({
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  children: [new Paragraph({ text: "Defect ID" })],
                  width: {
                    size: 15,
                    type: WidthType.PERCENTAGE,
                  },
                }),
                new TableCell({
                  children: [new Paragraph({ text: "Defect Type" })],
                  width: {
                    size: 20,
                    type: WidthType.PERCENTAGE,
                  },
                }),
                new TableCell({
                  children: [new Paragraph({ text: "Engineer's Remark" })],
                  width: {
                    size: 65,
                    type: WidthType.PERCENTAGE,
                  },
                }),
              ],
            }),
            ...imageDefects.map(defect => 
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph(defect.id)],
                  }),
                  new TableCell({
                    children: [new Paragraph(defect.type)],
                  }),
                  new TableCell({
                    children: [new Paragraph(defect.notes || "N.A")],
                  }),
                ],
              })
            ),
          ],
        }),
        new Paragraph({ text: "" }),
        new Paragraph({ text: "" })
      );
    }

    // Create document without summary section
    const doc = new Document({
      sections: [{
        properties: {},
        children: sections,
      }],
    });

    // Generate and save document
    const buffer = await Packer.toBlob(doc);
    saveAs(buffer, `inspection_report_${formData.location}.docx`);
  };

  const processImageToBase64 = async (file: File): Promise<Uint8Array> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        const img = new Image();
        
        img.onload = () => {
          canvas.width = img.width;
          canvas.height = img.height;
          ctx?.drawImage(img, 0, 0);
          const base64String = canvas.toDataURL('image/png').split(',')[1];
          resolve(Uint8Array.from(atob(base64String), c => c.charCodeAt(0)));
        };
        
        img.src = reader.result as string;
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  };

  const handleLocationPlanUpload = (file: File, imageIndex: number) => {
    const newImages = [...images];
    if (newImages[imageIndex]) {
      const img = new Image();
      img.src = URL.createObjectURL(file);
      newImages[imageIndex].locationPlan = file;
      newImages[imageIndex].locationPlanImg = img;
      newImages[imageIndex].locationPlanAnnotations = [];
      setImages(newImages);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-white shadow-sm border-b">
        <div className="container mx-auto px-4 py-4 flex justify-between items-center">
          <h1 className="text-2xl font-bold text-gray-900">AI Defect Detection and Report Generation</h1>
          <img 
            src="/prudential-logo.svg" 
            alt="Prudential Logo" 
            className="h-14"
          />
        </div>
      </header>
      
      <main className="container mx-auto px-4 py-6">
        <div className="space-y-6">
          {/* Image Upload and Annotation Section */}
          <div className="bg-white rounded-lg shadow-sm border p-4">
            <h2 className="text-lg font-semibold text-gray-700 mb-4">Image Upload and Annotation</h2>
            
            {/* Image Upload Section */}
            <div className="mb-6">
              <div className="bg-white rounded-lg p-4">
                <div className="space-y-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      Visual Image
                    </label>
                    <input
                      type="file"
                      accept="image/*"
                      multiple
                      onChange={(e) => handleImageUpload(e, 'visual')}
                      className="block w-full text-sm text-gray-500 file:mr-4 file:py-2.5 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-medium file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 transition-all"
                    />
                    {images.length > 0 && (
                      <p className="mt-2 text-sm text-gray-500">
                        {images.length} visual image(s) uploaded
                      </p>
                    )}
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      Infrared Image
                    </label>
                    <input
                      type="file"
                      accept="image/*"
                      multiple
                      onChange={(e) => handleImageUpload(e, 'infrared')}
                      className="block w-full text-sm text-gray-500 file:mr-4 file:py-2.5 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-medium file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 transition-all"
                    />
                    {parallelImages.infrared.length > 0 && (
                      <p className="mt-2 text-sm text-gray-500">
                        {parallelImages.infrared.length} infrared image(s) uploaded
                      </p>
                    )}
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      Hyperspectral Image
                    </label>
                    <input
                      type="file"
                      accept="image/*"
                      multiple
                      onChange={(e) => handleImageUpload(e, 'hyperspectral')}
                      className="block w-full text-sm text-gray-500 file:mr-4 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-medium file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 transition-all"
                    />
                    {parallelImages.hyperspectral.length > 0 && (
                      <p className="mt-2 text-sm text-gray-500">
                        {parallelImages.hyperspectral.length} hyperspectral image(s) uploaded
                      </p>
                    )}
                  </div>
                </div>
              </div>
            </div>

            {images.length > 0 && (
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div className="lg:col-span-2 space-y-6">
                  <div className="bg-white rounded-lg p-4">
                    <div className="flex items-center justify-between mb-4">
                      <button
                        onClick={goToPreviousImage}
                        disabled={currentImageIndex === 0}
                        className={`px-4 py-2 rounded-md transition-all ${
                          currentImageIndex === 0
                            ? 'bg-gray-100 text-gray-400 cursor-not-allowed'
                            : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-50'
                        }`}
                      >
                        ← Previous
                      </button>
                      <span className="text-sm font-medium text-gray-600">
                        Image {currentImageIndex + 1} of {images.length}
                      </span>
                      <button
                        onClick={goToNextImage}
                        disabled={currentImageIndex === images.length - 1}
                        className={`px-4 py-2 rounded-md transition-all ${
                          currentImageIndex === images.length - 1
                            ? 'bg-gray-100 text-gray-400 cursor-not-allowed'
                            : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-50'
                        }`}
                      >
                        Next →
                      </button>
                    </div>

                    {/* Main Image and Location Plan Display */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                      <div>
                        <h3 className="text-sm font-medium text-gray-700 mb-2">Visual Image</h3>
                        <Canvas
                          image={currentImage}
                          annotations={currentAnnotations}
                          setAnnotations={(newAnnotations) => {
                            const newImages = [...images];
                            if (newImages[currentImageIndex]) {
                              newImages[currentImageIndex].annotations = newAnnotations;
                              setImages(newImages);
                            }
                          }}
                          onSelectAnnotation={setSelectedAnnotationIndex}
                          onDimensionsChange={setCanvasDimensions}
                          onLocationPlanUpload={(file) => handleLocationPlanUpload(file, currentImageIndex)}
                        />
                      </div>
                      <div>
                        <h3 className="text-sm font-medium text-gray-700 mb-2">Location Plan</h3>
                        {images[currentImageIndex]?.locationPlanImg ? (
                          <Canvas
                            image={images[currentImageIndex].locationPlanImg}
                            annotations={images[currentImageIndex].locationPlanAnnotations || []}
                            setAnnotations={(newAnnotations) => {
                              const newImages = [...images];
                              if (newImages[currentImageIndex]) {
                                newImages[currentImageIndex].locationPlanAnnotations = newAnnotations;
                                setImages(newImages);
                              }
                            }}
                            onSelectAnnotation={() => {}}
                            onDimensionsChange={() => {}}
                            isLocationPlan={true}
                          />
                        ) : (
                          <div className="border-2 border-dashed border-gray-300 rounded-lg h-[400px] flex items-center justify-center">
                            <p className="text-gray-500">No location plan uploaded</p>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                </div>

                {selectedAnnotationIndex !== null && (
                  <div className="space-y-6">
                    <div className="bg-white rounded-lg p-4">
                      <h2 className="text-lg font-semibold text-gray-700 mb-4">Annotation Details</h2>
                      <div className="space-y-4">
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-1">
                            Defect Type
                          </label>
                          <select
                            value={currentAnnotations[selectedAnnotationIndex]?.defectType || ''}
                            onChange={(e) => handleAnnotationUpdate(selectedAnnotationIndex, { defectType: e.target.value })}
                            className="w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
                          >
                            <option value="">Select a defect type</option>
                            {defectTypes.map((type) => (
                              <option key={type} value={type}>
                                {type}
                              </option>
                            ))}
                          </select>
                        </div>
                        
                        <div className="space-y-2">
                          <label className="block text-sm font-medium text-gray-700">
                            Engineer's Remark
                          </label>
                          <textarea
                            value={currentAnnotations[selectedAnnotationIndex]?.notes || ''}
                            onChange={(e) => handleAnnotationUpdate(selectedAnnotationIndex, { notes: e.target.value })}
                            placeholder="Enter engineer's remark here..."
                            className="w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-sm"
                            rows={3}
                          />
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>

        {images.length > 0 && (
          <div className="fixed bottom-0 left-0 right-0 bg-white border-t shadow-lg">
            <div className="container mx-auto px-4 py-4">
              <div className="flex justify-end gap-4">
                <button
                  onClick={saveCurrentImage}
                  className="inline-flex items-center px-4 py-2 border border-gray-300 rounded-md shadow-sm text-sm font-medium text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-all"
                >
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2 text-gray-500" viewBox="0 0 20 20" fill="currentColor">
                    <path d="M7.707 10.293a1 1 0 10-1.414 1.414l3 3a1 1 0 001.414 0l3-3a1 1 0 00-1.414-1.414L11 11.586V6h-2v5.586l-1.293-1.293z" />
                    <path d="M4 4a2 2 0 012-2h8a2 2 0 012 2v12a2 2 0 01-2 2H6a2 2 0 01-2-2V4zm2-1a1 1 0 00-1 1v12a1 1 0 001 1h8a1 1 0 001-1V4a1 1 0 00-1-1H6z" />
                  </svg>
                  Save Current Image
                </button>
                <button
                  id="save-all-button"
                  onClick={saveAllImages}
                  className="inline-flex items-center px-4 py-2 border border-gray-300 rounded-md shadow-sm text-sm font-medium text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-all disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2 text-gray-500" viewBox="0 0 20 20" fill="currentColor">
                    <path fillRule="evenodd" d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zm3.293-7.707a1 1 0 011.414 0L9 10.586V3a1 1 0 112 0v7.586l1.293-1.293a1 1 0 111.414 1.414l-3 3a1 1 0 01-1.414 0l-3-3a1 1 0 010-1.414z" clipRule="evenodd" />
                  </svg>
                  Save All Images
                </button>
                <button
                  onClick={generateReport}
                  className="inline-flex items-center px-4 py-2 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-all"
                >
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2" viewBox="0 0 20 20" fill="currentColor">
                    <path fillRule="evenodd" d="M4 4a2 2 0 012-2h4.586A2 2 0 0112 2.586L15.414 6A2 2 0 0116 7.414V16a2 2 0 01-2 2H6a2 2 0 01-2-2V4zm2 6a1 1 0 011-1h6a1 1 0 100 2H7a1 1 0 01-1-1zm1 3a1 1 0 100 2h6a1 1 0 100-2H7z" clipRule="evenodd" />
                  </svg>
                  Generate Report
                </button>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
