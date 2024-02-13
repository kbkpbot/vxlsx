
Need to fix some .h, "zlib.h" ==> <zlib.h>

libxlsxwriter-in.c :
```c
/**
 * \file libxlsxwriter.c
 * Single-file libxlsxwriter library.
 *
 * You must install minizip lib :
 * apt install libminizip-dev
 *
 * And link with minizip
 * -lminizip
 *
 * Generate using:
 * \code
 *      python combine.py -r ../include -r ../third_party -r .. -o libxlsxwriter.c libxlsxwriter-in.c
 * \endcode
 */
#define USE_SYSTEM_MINIZIP
#define LXW_HAS_SNPRINTF

#include "xlsxwriter.h"

#include "../src/vml.c"
#include "../src/chartsheet.c"
#include "../src/theme.c"
#include "../src/content_types.c"
#include "../src/xmlwriter.c"
#include "../src/app.c"
#include "../src/styles.c"
#include "../src/core.c"
#include "../src/comment.c"
#include "../src/utility.c"
#include "../src/metadata.c"
#include "../src/custom.c"
#include "../src/hash_table.c"
#include "../src/relationships.c"
#include "../src/drawing.c"
#include "../src/chart.c"
#include "../src/shared_strings.c"
#include "../src/worksheet.c"
#include "../src/format.c"
#include "../src/table.c"
#include "../src/workbook.c"
#include "../src/packager.c"
#include "../third_party/tmpfileplus/tmpfileplus.c"
#include "../third_party/md5/md5.c"

//#include "../third_party/minizip/minizip.c"
//#include "../third_party/minizip/ioapi.c"
//#include "../third_party/minizip/zip.c"
//#include "../third_party/minizip/crc32.c"
//#include "../third_party/minizip/ioapi.c"

```

# TODO
Use integrated minizip, not the system minizip