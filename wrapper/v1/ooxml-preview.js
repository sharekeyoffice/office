(function (global) {
  "use strict";

  const namespaces = {
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
    p: "http://schemas.openxmlformats.org/presentationml/2006/main",
    xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    s: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    rel: "http://schemas.openxmlformats.org/package/2006/relationships",
    mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
    ct: "http://schemas.openxmlformats.org/package/2006/content-types",
    wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    a14: "http://schemas.microsoft.com/office/drawing/2010/main"
  };

  const zipLoadPromises = new WeakMap();

  function logSkippedPreview(reason, context) {
    console.warn("[office-preview] " + reason, context);
  }

  function renderPreview(graphicObject, context) {
    const canRenderPreview = graphicObject && (
      (graphicObject.isShape && graphicObject.isShape()) ||
      (graphicObject.isChart && graphicObject.isChart())
    );

    if (!canRenderPreview) {
      if (!graphicObject) {
        logSkippedPreview("Missing graphic object", context);
      }

      return null;
    }

    try {
      const imageDataUrl = graphicObject.getBase64Img(true, "image/png");
      const isPngImage = /^data:image\/png;base64,/.test(imageDataUrl);

      if (!isPngImage) {
        logSkippedPreview("Preview renderer returned a non-PNG image", context);

        return null;
      }

      const drawingBounds = graphicObject.getBoundsByDrawing
        ? graphicObject.getBoundsByDrawing(true)
        : null;

      return {
        data: imageDataUrl.slice(imageDataUrl.indexOf(",") + 1),
        bounds: drawingBounds ? {
          x: drawingBounds.l,
          y: drawingBounds.t,
          width: drawingBounds.r - drawingBounds.l,
          height: drawingBounds.b - drawingBounds.t
        } : null
      };
    } catch (error) {
      logSkippedPreview("Object preview could not be rendered", { ...context, error });

      return null;
    }
  }

  function capture(win, format) {
    format = format || "docx";

    const editorApi = win.Asc && win.Asc.editor;
    const documentModel = editorApi &&
      editorApi.WordControl &&
      editorApi.WordControl.m_oLogicDocument;
    const images = [];
    const canCaptureSpreadsheet = format === "xlsx" && editorApi && editorApi.wb;

    if (canCaptureSpreadsheet) {
      editorApi.wb.model.aWorksheets.forEach(function (worksheet, sheetIndex) {
        worksheet.Drawings.forEach(function (drawing, drawingIndex) {
          const graphicObject = drawing.graphicObject;
          const nonVisualProperties = graphicObject.getUniNvProps &&
            graphicObject.getUniNvProps();
          const drawingProperties = nonVisualProperties && nonVisualProperties.cNvPr;
          const previewContext = {
            format,
            sheetIndex,
            drawingIndex,
            drawingId: drawingProperties ? drawingProperties.id : null
          };
          const isHiddenOrMissingProperties = !drawingProperties || drawingProperties.isHidden;

          if (isHiddenOrMissingProperties) {
            if (!drawingProperties) {
              logSkippedPreview("Missing drawing properties", previewContext);
            }

            return;
          }

          const image = renderPreview(graphicObject, previewContext);
          const hasValidBounds = image &&
            image.bounds &&
            Object.values(image.bounds).every(Number.isFinite);

          if (!hasValidBounds) {
            if (image) {
              logSkippedPreview("Preview bounds are missing or invalid", previewContext);
            }

            return;
          }

          image.id = String(drawingProperties.id);
          image.sheet = worksheet.sName;
          images.push(image);
        });
      });

      return { format, images };
    }

    if (!documentModel) {
      logSkippedPreview("Editor model is unavailable", { format });

      return { format, images };
    }

    const canCapturePresentation = format === "pptx" && documentModel.Slides;
    const canCaptureWordDocument = format === "docx" &&
      documentModel.GetAllDrawingObjects;

    if (canCapturePresentation) {
      documentModel.Slides.forEach(function (slide, slideIndex) {
        slide.cSld.spTree.forEach(function (graphicObject, drawingIndex) {
          const nonVisualProperties = graphicObject.getUniNvProps &&
            graphicObject.getUniNvProps();
          const drawingProperties = nonVisualProperties && nonVisualProperties.cNvPr;
          const previewContext = {
            format,
            drawingIndex,
            drawingId: drawingProperties ? drawingProperties.id : null,
            part: "ppt/slides/slide" + (slideIndex + 1) + ".xml"
          };
          const shouldSkipDrawing = !drawingProperties ||
            drawingProperties.isHidden ||
            (graphicObject.isPlaceholder && graphicObject.isPlaceholder());

          if (shouldSkipDrawing) {
            if (!drawingProperties) {
              logSkippedPreview("Missing drawing properties", previewContext);
            }

            return;
          }

          const image = renderPreview(graphicObject, previewContext);
          const hasValidBounds = image &&
            image.bounds &&
            Object.values(image.bounds).every(Number.isFinite);

          if (!hasValidBounds) {
            if (image) {
              logSkippedPreview("Preview bounds are missing or invalid", previewContext);
            }

            return;
          }

          image.id = String(drawingProperties.id);
          image.part = "ppt/slides/slide" + (slideIndex + 1) + ".xml";
          images.push(image);
        });
      });
    } else if (canCaptureWordDocument) {
      documentModel.GetAllDrawingObjects().forEach(function (drawing, drawingIndex) {
        const previewContext = {
          format,
          drawingIndex,
          drawingId: drawing.docPr ? drawing.docPr.id : null
        };
        const shouldSkipDrawing = !drawing.docPr || drawing.docPr.isHidden;

        if (shouldSkipDrawing) {
          if (!drawing.docPr) {
            logSkippedPreview("Missing drawing properties", previewContext);
          }

          return;
        }

        const image = renderPreview(drawing.GraphicObj, previewContext);

        if (!image) {
          return;
        }

        image.id = String(drawing.docPr.id);
        images.push(image);
      });
    }

    return { format, images };
  }

  function loadZip(win) {
    if (win.AscCommon.ZLib) {
      return Promise.resolve(win.AscCommon.ZLib);
    }

    if (!zipLoadPromises.has(win)) {
      const loadPromise = new Promise(function (resolve, reject) {
        const timeout = setTimeout(function () {
          reject(new Error("Office preview ZIP engine timed out"));
        }, 15e3);

        win.AscCommon.zlib_load(
          function () {
            clearTimeout(timeout);
            resolve(win.AscCommon.ZLib);
          },
          function () {
            clearTimeout(timeout);
            reject(new Error("Office preview ZIP engine failed to load"));
          }
        );
      }).catch(function (error) {
        zipLoadPromises.delete(win);

        throw error;
      });

      zipLoadPromises.set(win, loadPromise);
    }

    return zipLoadPromises.get(win);
  }

  function parseXml(bytes) {
    const xmlText = new TextDecoder().decode(bytes);
    const xmlDocument = new DOMParser().parseFromString(xmlText, "application/xml");

    if (xmlDocument.getElementsByTagName("parsererror").length) {
      throw new Error("Invalid OOXML XML");
    }

    return xmlDocument;
  }

  function writeZipFile(archive, filePath, bytes) {
    if (!archive.addFile(filePath, bytes)) {
      throw new Error("Could not update OOXML part: " + filePath);
    }
  }

  function serializeXml(xmlDocument) {
    const xmlText = new XMLSerializer().serializeToString(xmlDocument);

    return new TextEncoder().encode(xmlText);
  }

  function createElement(xmlDocument, namespacePrefix, localName, attributes) {
    const element = xmlDocument.createElementNS(
      namespaces[namespacePrefix],
      namespacePrefix + ":" + localName
    );

    Object.keys(attributes || {}).forEach(function (attributeName) {
      element.setAttribute(attributeName, attributes[attributeName]);
    });

    return element;
  }

  function createPicture(xmlDocument, namespacePrefix, relationshipId, frame) {
    const picture = createElement(xmlDocument, namespacePrefix, "pic");
    const nonVisualProperties = createElement(xmlDocument, namespacePrefix, "nvPicPr");

    nonVisualProperties.append(
      createElement(xmlDocument, namespacePrefix, "cNvPr", {
        id: "0",
        name: "Object preview"
      }),
      createElement(xmlDocument, namespacePrefix, "cNvPicPr")
    );

    if (namespacePrefix === "p") {
      nonVisualProperties.append(createElement(xmlDocument, "p", "nvPr"));
    }

    const imageFill = createElement(xmlDocument, namespacePrefix, "blipFill");
    const imageReference = createElement(xmlDocument, "a", "blip");
    imageReference.setAttributeNS(namespaces.r, "r:embed", relationshipId);

    const stretch = createElement(xmlDocument, "a", "stretch");
    stretch.append(createElement(xmlDocument, "a", "fillRect"));
    imageFill.append(imageReference, stretch);

    const shapeProperties = createElement(xmlDocument, namespacePrefix, "spPr");
    const transform = createElement(xmlDocument, "a", "xfrm");

    transform.append(
      createElement(xmlDocument, "a", "off", {
        x: String(frame.x),
        y: String(frame.y)
      }),
      createElement(xmlDocument, "a", "ext", {
        cx: String(frame.width),
        cy: String(frame.height)
      })
    );

    const geometry = createElement(xmlDocument, "a", "prstGeom", { prst: "rect" });
    geometry.append(createElement(xmlDocument, "a", "avLst"));
    shapeProperties.append(transform, geometry);
    picture.append(nonVisualProperties, imageFill, shapeProperties);

    return picture;
  }

  function createWordPreview(xmlDocument, drawing, relationshipId) {
    const drawingCopy = drawing.cloneNode(true);
    const graphicData = drawingCopy.getElementsByTagNameNS(namespaces.a, "graphicData")[0];
    const extent = drawingCopy.getElementsByTagNameNS(namespaces.wp, "extent")[0];

    if (!graphicData || !extent) {
      return null;
    }

    graphicData.replaceChildren(createPicture(xmlDocument, "pic", relationshipId, {
      x: 0,
      y: 0,
      width: extent.getAttribute("cx"),
      height: extent.getAttribute("cy")
    }));
    graphicData.setAttribute("uri", namespaces.pic);

    return drawingCopy;
  }

  function resolvePartPath(basePath, targetPath) {
    let pathSegments = basePath.split("/");
    pathSegments.pop();

    if (targetPath[0] === "/") {
      pathSegments = [];
    }

    targetPath.split("/").forEach(function (segment) {
      if (segment === "..") {
        pathSegments.pop();
      } else if (segment && segment !== ".") {
        pathSegments.push(segment);
      }
    });

    return pathSegments.join("/");
  }

  function getSpreadsheetDrawingSheets(archive) {
    const drawingSheetsByPart = {};
    const workbookDocument = parseXml(archive.getFile("xl/workbook.xml"));
    const workbookRelationships = parseXml(
      archive.getFile("xl/_rels/workbook.xml.rels")
    );
    const worksheets = Array.from(
      workbookDocument.getElementsByTagNameNS(namespaces.s, "sheet")
    );

    for (const worksheet of worksheets) {
      const worksheetRelationship = Array.from(
        workbookRelationships.documentElement.children
      ).find(function (relationship) {
        return relationship.getAttribute("Id") ===
          worksheet.getAttributeNS(namespaces.r, "id");
      });

      if (!worksheetRelationship) {
        logSkippedPreview("Worksheet relationship is missing", {
          format: "xlsx",
          relationshipId: worksheet.getAttributeNS(namespaces.r, "id")
        });

        continue;
      }

      const worksheetPath = resolvePartPath(
        "xl/workbook.xml",
        worksheetRelationship.getAttribute("Target")
      );
      const relationshipsPath = worksheetPath.replace(/([^/]+)$/, "_rels/$1.rels");
      const relationshipsBytes = archive.getFile(relationshipsPath);

      if (!relationshipsBytes) {
        continue;
      }

      const worksheetRelationships = parseXml(relationshipsBytes);

      for (const relationship of Array.from(worksheetRelationships.documentElement.children)) {
        const isDrawingRelationship = relationship.getAttribute("Type") ===
          namespaces.r + "/drawing";

        if (isDrawingRelationship) {
          const drawingPartPath = resolvePartPath(
            worksheetPath,
            relationship.getAttribute("Target")
          );

          drawingSheetsByPart[drawingPartPath] = worksheet.getAttribute("name");
        }
      }
    }

    return drawingSheetsByPart;
  }

  function getDrawingTargets(xmlDocument, isPresentation, isSpreadsheet) {
    if (isSpreadsheet) {
      const shapes = Array.from(xmlDocument.getElementsByTagNameNS(namespaces.xdr, "sp"));
      const graphicFrames = Array.from(
        xmlDocument.getElementsByTagNameNS(namespaces.xdr, "graphicFrame")
      );

      return shapes.concat(graphicFrames).filter(function (drawing) {
        return /^(twoCellAnchor|oneCellAnchor|absoluteAnchor)$/.test(
          drawing.parentNode.localName
        );
      });
    }

    if (isPresentation) {
      const shapeTree = xmlDocument.getElementsByTagNameNS(namespaces.p, "spTree")[0];

      return Array.from(shapeTree.children).filter(function (drawing) {
        const isSupportedDrawing = drawing.namespaceURI === namespaces.p && (
          drawing.localName === "sp" || drawing.localName === "graphicFrame"
        );

        return isSupportedDrawing;
      });
    }

    return Array.from(xmlDocument.getElementsByTagNameNS(namespaces.w, "drawing"));
  }

  async function apply(bytes, snapshot, win) {
    if (!snapshot || !snapshot.images.length) {
      return bytes;
    }

    const previewImages = snapshot.images;
    const isPresentation = snapshot.format === "pptx";
    const isSpreadsheet = snapshot.format === "xlsx";
    const isWordDocument = !isPresentation && !isSpreadsheet;
    const drawingNamespacePrefix = isPresentation ? "p" : isSpreadsheet ? "xdr" : "wp";
    const drawingPropertyName = drawingNamespacePrefix === "wp" ? "docPr" : "cNvPr";
    const ZipArchive = await loadZip(win);
    const archive = new ZipArchive();

    try {
      if (!archive.open(bytes)) {
        throw new Error("Could not open OOXML for compatibility");
      }

      const contentTypesDocument = parseXml(archive.getFile("[Content_Types].xml"));
      const drawingSheetsByPart = isSpreadsheet
        ? getSpreadsheetDrawingSheets(archive)
        : {};
      const drawingPartPattern = isSpreadsheet
        ? /^xl\/drawings\/drawing\d+\.xml$/
        : isPresentation
          ? /^ppt\/slides\/slide\d+\.xml$/
          : /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;
      const drawingPartPaths = archive.getPaths().filter(function (filePath) {
        return drawingPartPattern.test(filePath);
      });
      const imageDirectory = isPresentation
        ? "ppt/media/"
        : isSpreadsheet ? "xl/media/" : "word/media/";

      let archiveChanged = false;
      let previewIndex = 0;
      let nextDrawingId = 0;

      for (const partPath of drawingPartPaths) {
        const partDocument = parseXml(archive.getFile(partPath));
        const drawingProperties = Array.from(
          partDocument.getElementsByTagNameNS(
            namespaces[drawingNamespacePrefix],
            drawingPropertyName
          )
        );

        for (const drawingProperty of drawingProperties) {
          const drawingId = Number(drawingProperty.getAttribute("id")) || 0;
          nextDrawingId = Math.max(nextDrawingId, drawingId);
        }
      }

      for (const partPath of drawingPartPaths) {
        const partDocument = parseXml(archive.getFile(partPath));
        const relationshipsPath = partPath.replace(/([^/]+)$/, "_rels/$1.rels");
        const relationshipsBytes = archive.getFile(relationshipsPath);
        const relationshipsDocument = relationshipsBytes
          ? parseXml(relationshipsBytes)
          : new DOMParser().parseFromString(
            '<Relationships xmlns="' + namespaces.rel + '"/>',
            "application/xml"
          );
        const drawings = getDrawingTargets(partDocument, isPresentation, isSpreadsheet);

        let partChanged = false;

        for (const drawing of drawings) {
          let insertionNode = isSpreadsheet ? drawing.parentNode : drawing;

          if (isWordDocument) {
            const parentNode = drawing.parentNode;
            const isAlternateContentChoice = parentNode.namespaceURI === namespaces.mc &&
              parentNode.localName === "Choice";

            if (isAlternateContentChoice) {
              insertionNode = parentNode.parentNode;
            }

            const isInsideWordRun = insertionNode.parentNode.namespaceURI === namespaces.w &&
              insertionNode.parentNode.localName === "r";

            if (!isInsideWordRun) {
              continue;
            }

            const isSupportedWordDrawing = drawing.getElementsByTagNameNS(
              namespaces.wps,
              "wsp"
            ).length || drawing.getElementsByTagNameNS(namespaces.c, "chart").length;

            if (!isSupportedWordDrawing) {
              continue;
            }
          }

          const drawingProperties = drawing.getElementsByTagNameNS(
            namespaces[drawingNamespacePrefix],
            drawingPropertyName
          )[0];
          const matchingImages = drawingProperties && previewImages.filter(function (image) {
            const matchesDrawingId = image.id === drawingProperties.getAttribute("id");
            const matchesPart = !image.part || image.part === partPath;
            const matchesWorksheet = !isSpreadsheet ||
              image.sheet === drawingSheetsByPart[partPath];

            return matchesDrawingId && matchesPart && matchesWorksheet;
          });
          const hasUniquePreview = matchingImages && matchingImages.length === 1;

          if (!hasUniquePreview) {
            if (!drawingProperties) {
              logSkippedPreview("OOXML drawing properties are missing", {
                format: snapshot.format,
                part: partPath
              });
            } else if (matchingImages.length > 1) {
              logSkippedPreview("Multiple previews match the same drawing", {
                format: snapshot.format,
                part: partPath,
                drawingId: drawingProperties.getAttribute("id"),
                matches: matchingImages.length
              });
            }

            continue;
          }

          const image = matchingImages[0];
          let imagePath;
          let relationshipId;
          let hasPreviewNameConflict;

          do {
            ++previewIndex;
            imagePath = imageDirectory + "sk-shape-preview-" + previewIndex + ".png";
            relationshipId = "skShapePreview" + previewIndex;

            const imagePathExists = archive.getPaths().includes(imagePath);
            const relationshipIdExists = Array.from(
              relationshipsDocument.documentElement.children
            ).some(function (relationship) {
              return relationship.getAttribute("Id") === relationshipId;
            });

            hasPreviewNameConflict = imagePathExists || relationshipIdExists;
          } while (hasPreviewNameConflict);

          let preview;

          if (isPresentation || isSpreadsheet) {
            const bounds = image.bounds;

            preview = createPicture(partDocument, drawingNamespacePrefix, relationshipId, {
              x: Math.round(bounds.x * 36e3),
              y: Math.round(bounds.y * 36e3),
              width: Math.round(bounds.width * 36e3),
              height: Math.round(bounds.height * 36e3)
            });
          } else {
            preview = createWordPreview(partDocument, drawing, relationshipId);
          }

          if (!preview) {
            logSkippedPreview("Word drawing graphicData or extent is missing", {
              format: snapshot.format,
              part: partPath,
              drawingId: image.id
            });

            continue;
          }

          const previewProperties = preview.getElementsByTagNameNS(
            namespaces[drawingNamespacePrefix],
            drawingPropertyName
          )[0];
          previewProperties.setAttribute("id", String(++nextDrawingId));
          previewProperties.setAttribute("name", "SK object preview " + previewIndex);

          if (isSpreadsheet) {
            // Use an absolute anchor for Quick Look; keep cell anchoring on the editable shape.
            const anchor = createElement(partDocument, "xdr", "absoluteAnchor");

            anchor.append(
              createElement(partDocument, "xdr", "pos", {
                x: Math.round(image.bounds.x * 36e3),
                y: Math.round(image.bounds.y * 36e3)
              }),
              createElement(partDocument, "xdr", "ext", {
                cx: Math.round(image.bounds.width * 36e3),
                cy: Math.round(image.bounds.height * 36e3)
              }),
              preview,
              createElement(partDocument, "xdr", "clientData")
            );
            preview = anchor;
          }

          // Keep the editable object in Choice and expose the bitmap only as its fallback.
          const canReuseAlternateContent = isWordDocument &&
            insertionNode.namespaceURI === namespaces.mc &&
            insertionNode.localName === "AlternateContent";

          if (canReuseAlternateContent) {
            let fallback = Array.from(insertionNode.children).find(function (childNode) {
              return childNode.namespaceURI === namespaces.mc &&
                childNode.localName === "Fallback";
            });

            if (!fallback) {
              fallback = createElement(partDocument, "mc", "Fallback");
              insertionNode.append(fallback);
            }

            fallback.replaceChildren(preview);
          } else {
            const alternateContent = createElement(partDocument, "mc", "AlternateContent");
            alternateContent.setAttributeNS(
              "http://www.w3.org/2000/xmlns/",
              "xmlns:a14",
              namespaces.a14
            );

            const choice = createElement(partDocument, "mc", "Choice", { Requires: "a14" });
            const fallback = createElement(partDocument, "mc", "Fallback");

            insertionNode.parentNode.insertBefore(alternateContent, insertionNode);
            choice.append(insertionNode);
            fallback.append(preview);
            alternateContent.append(choice, fallback);
          }

          const decodedImage = atob(image.data);
          const pngBytes = Uint8Array.from(decodedImage, function (character) {
            return character.charCodeAt(0);
          });
          writeZipFile(archive, imagePath, pngBytes);

          const relationship = relationshipsDocument.createElementNS(
            namespaces.rel,
            "Relationship"
          );
          const relationshipTarget = isPresentation || isSpreadsheet
            ? "../media/" + imagePath.split("/").pop()
            : imagePath.slice(5);

          relationship.setAttribute("Id", relationshipId);
          relationship.setAttribute("Type", namespaces.r + "/image");
          relationship.setAttribute("Target", relationshipTarget);
          relationshipsDocument.documentElement.append(relationship);
          partChanged = true;
        }

        if (partChanged) {
          writeZipFile(archive, partPath, serializeXml(partDocument));
          writeZipFile(archive, relationshipsPath, serializeXml(relationshipsDocument));
          archiveChanged = true;
        }
      }

      if (!archiveChanged) {
        return bytes;
      }

      const hasPngContentType = Array.from(
        contentTypesDocument.getElementsByTagNameNS(namespaces.ct, "Default")
      ).some(function (contentType) {
        return contentType.getAttribute("Extension").toLowerCase() === "png";
      });

      if (!hasPngContentType) {
        const pngContentType = contentTypesDocument.createElementNS(namespaces.ct, "Default");
        pngContentType.setAttribute("Extension", "png");
        pngContentType.setAttribute("ContentType", "image/png");
        contentTypesDocument.documentElement.append(pngContentType);
        writeZipFile(archive, "[Content_Types].xml", serializeXml(contentTypesDocument));
      }

      const outputBytes = archive.save();

      if (!outputBytes) {
        throw new Error("Could not save OOXML previews");
      }

      // Copy out of WASM memory before closing the archive.
      return Uint8Array.from(outputBytes);
    } finally {
      archive.close();
    }
  }

  global.OfficePreview = { capture, apply };
})(window);
