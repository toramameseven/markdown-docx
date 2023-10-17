export class Bookmarks {
  bookmarkMap = new Map();
  constructor() {}

  clear() {
    this.bookmarkMap = new Map();
  }

  has(testId: string) {
    return this.bookmarkMap.has(testId);
  }

  // https://qiita.com/satokaz/items/64582da4640898c4bf42
  // slugify:
  slugify(header: string, alowDuplicate = false) {
    //return encodeURI(
    let r = header
      .trim()
      .toLowerCase()
      .replace(
        /[\]\[\!\"\#\$\%\&\'\(\)\*\+\,\.\/\:\;\<\=\>\?\@\\\^\_\{\|\}\~＠＃＄％＾＆＊（）＿＋－＝｛｝”’＜＞［］「」・、。～]/g,
        ""
      )
      .replace(/\s+/g, "-") // Replace spaces with hyphens
      .replace(/\-+$/, ""); // Replace trailing hyphen

    if (alowDuplicate === false) {
      r = this.createUniqId(r, r);
    }

    return r;
  }

  private createUniqId(id: string, originalId: string, index = 0) {
    let testId = id;
    if (this.bookmarkMap.has(testId)) {
      testId = this.createUniqId(
        originalId + "-" + (index + 1).toString(),
        originalId,
        index + 1
      );
    }
    this.bookmarkMap.set(testId, testId);
    return testId;
  }
}
