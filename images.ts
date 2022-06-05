function fitImage(
  image: GoogleAppsScript.Spreadsheet.OverGridImage, 
  maxWidth: number, 
  maxHeight: number,
) {
  const inherentWidth = image.getInherentWidth();
  const inherentHeight = image.getInherentHeight();

  let width = inherentWidth
  let height = inherentHeight

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
