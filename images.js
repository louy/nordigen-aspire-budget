/** @param image {SpreadsheetApp.OverGridImage} */
function fitImage(image, maxWidth, maxHeight) {
  const inherentWidth = image.getInherentWidth();
  const inherentHeight = image.getInherentHeight();

  let width = inherentWidth
  let height = inherentHeight

  const ratio = width / height

  if (width > maxWidth) {
    width = maxWidth
    height = inherentHeight * maxWidth / inherentWidth;
  }

  if (height > maxHeight) {
    height = maxHeight
    width = inherentWidth * maxHeight / inherentHeight
  }

  image.setWidth(width)
  image.setHeight(height)
}
