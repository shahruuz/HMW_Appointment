const SPREADSHEET_ID =
  '18EdJRs8KM2DzAistPis1Hc-iitagmAQYPNEr9ccmEX0';

const SHEET_NAME =
  'Website_Requests';


function doGet() {

  return HtmlService
    .createHtmlOutputFromFile(
      'Index'
    )
    .setTitle(
      'Request an Appointment | Her Mind & Wellness'
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


function submitForm(data) {

  let cache = null;

  let cacheKey = '';

  let markedProcessing =
    false;

  let writeSucceeded =
    false;


  try {

    if (!data) {

      throw new Error(
        'No form information was received.'
      );

    }


    const submissionId =
      clean_(
        data.submissionId
      );


    const firstName =
      clean_(
        data.firstName
      );


    const lastName =
      clean_(
        data.lastName
      );


    const phone =
      clean_(
        data.phone
      );


    const email =
      clean_(
        data.email
      )
      .toLowerCase();


    const zipCode =
      normalizeZip_(
        data.zipCode
      );


    const paymentType =
      clean_(
        data.paymentType
      );


    const visitPreference =
      clean_(
        data.visitPreference
      );


    const website =
      clean_(
        data.website
      );


    /* Honeypot */

    if (website) {

      return {
        success:
          false,

        message:
          'Unable to submit request.'
      };

    }


    /* Submission ID */

    if (
      !/^[A-Za-z0-9_-]{8,100}$/
        .test(
          submissionId
        )
    ) {

      throw new Error(
        'Unable to identify this request. Please refresh and try again.'
      );

    }


    /* Required */

    if (!firstName) {

      throw new Error(
        'Please enter your first name.'
      );

    }


    if (!lastName) {

      throw new Error(
        'Please enter your last name.'
      );

    }


    if (!phone) {

      throw new Error(
        'Please enter your phone number. Example: 915-555-1234.'
      );

    }


    if (!email) {

      throw new Error(
        'Please enter your email address. Example: name@example.com.'
      );

    }


    if (!zipCode) {

      throw new Error(
        'Please enter your ZIP code. Example: 79912.'
      );

    }


    if (!paymentType) {

      throw new Error(
        'Please choose how you would like to get started.'
      );

    }


    if (!visitPreference) {

      throw new Error(
        'Please choose a visit preference.'
      );

    }


    /* Validation */

    if (
      !isValidPhone_(
        phone
      )
    ) {

      throw new Error(
        'Please check your phone number. Example: 915-555-1234.'
      );

    }


    if (
      !isValidEmail_(
        email
      )
    ) {

      throw new Error(
        'Please check your email address. Example: name@example.com.'
      );

    }


    if (
      !isValidZip_(
        zipCode
      )
    ) {

      throw new Error(
        'Please enter a ZIP code like 79912 or 79912-1234.'
      );

    }


    const allowedPayments = [

      'Self-pay / Private pay',

      'Insurance',

      'Free 15-minute consultation'

    ];


    const allowedVisits = [

      'Telehealth - Texas',

      'In Person - El Paso'

    ];


    if (
      !allowedPayments.includes(
        paymentType
      )
    ) {

      throw new Error(
        'Please choose one of the available options.'
      );

    }


    if (
      !allowedVisits.includes(
        visitPreference
      )
    ) {

      throw new Error(
        'Please choose telehealth or in-person care.'
      );

    }



    /* --------------------------------------
       RETRY / DUPLICATE PROTECTION
       -------------------------------------- */

    cache =
      CacheService
        .getScriptCache();


    cacheKey =
      'HMW_' +
      submissionId;


    const existing =
      cache.get(
        cacheKey
      );


    if (
      existing ===
      'DONE'
    ) {

      return {

        success:
          true,

        paymentType:
          paymentType,

        duplicate:
          true

      };

    }


    if (
      existing ===
      'PROCESSING'
    ) {

      return {

        success:
          false,

        processing:
          true,

        retryable:
          true,

        message:
          'Your request is still processing. Please wait a few seconds.'

      };

    }


    try {

      cache.put(
        cacheKey,
        'PROCESSING',
        120
      );


      markedProcessing =
        true;

    }

    catch (cacheError) {}


    /* --------------------------------------
       ONE SHEET WRITE
       -------------------------------------- */

    const sheet =
      SpreadsheetApp
        .openById(
          SPREADSHEET_ID
        )
        .getSheetByName(
          SHEET_NAME
        );


    if (!sheet) {

      throw new Error(
        'Appointment request sheet was not found.'
      );

    }


    sheet.appendRow([

      new Date(),

      safeCell_(
        firstName
      ),

      safeCell_(
        lastName
      ),

      safeCell_(
        phone
      ),

      safeCell_(
        email
      ),

      safeCell_(
        zipCode
      ),

      safeCell_(
        paymentType
      ),

      safeCell_(
        visitPreference
      )

    ]);


    writeSucceeded =
      true;


    /*
     * Mark completed.
     */

    try {

      cache.put(
        cacheKey,
        'DONE',
        900
      );

    }

    catch (cacheError) {}


    return {

      success:
        true,

      paymentType:
        paymentType

    };

  }

  catch (error) {

    if (
      !writeSucceeded &&
      markedProcessing &&
      cache &&
      cacheKey
    ) {

      try {

        cache.remove(
          cacheKey
        );

      }

      catch (cacheError) {}

    }


    return {

      success:
        false,

      message:
        error &&
        error.message

          ? error.message

          : 'Something went wrong. Please try again.'

    };

  }

}



/* ------------------------------------------
   CHECK SLOW REQUEST
   ------------------------------------------ */

function checkSubmission(
  submissionId
) {

  const id =
    clean_(
      submissionId
    );


  if (
    !/^[A-Za-z0-9_-]{8,100}$/
      .test(id)
  ) {

    return {
      status:
        'unknown'
    };

  }


  try {

    const value =
      CacheService
        .getScriptCache()
        .get(
          'HMW_' +
          id
        );


    if (
      value ===
      'DONE'
    ) {

      return {
        status:
          'done'
      };

    }


    if (
      value ===
      'PROCESSING'
    ) {

      return {
        status:
          'processing'
      };

    }

  }

  catch (error) {}


  return {
    status:
      'unknown'
  };

}



/* ------------------------------------------
   HELPERS
   ------------------------------------------ */

function clean_(value) {

  if (
    value === null ||
    value === undefined
  ) {

    return '';

  }


  return String(value)
    .trim();

}


function normalizeZip_(value) {

  let zip =
    clean_(
      value
    )
    .replace(
      /\s+/g,
      ''
    );


  if (
    /^\d{9}$/.test(
      zip
    )
  ) {

    zip =
      zip.substring(
        0,
        5
      ) +
      '-' +
      zip.substring(
        5
      );

  }


  return zip;

}


function isValidPhone_(phone) {

  if (
    !/^\+?[0-9().\-\s]+$/
      .test(phone)
  ) {

    return false;

  }


  const digits =
    phone.replace(
      /\D/g,
      ''
    );


  return (
    digits.length >= 7 &&
    digits.length <= 15
  );

}


function isValidEmail_(email) {

  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/
    .test(email);

}


function isValidZip_(zip) {

  return /^\d{5}(?:-\d{4})?$/
    .test(zip);

}


function safeCell_(value) {

  const text =
    String(value);


  if (
    /^[=+\-@]/.test(
      text
    )
  ) {

    return "'" + text;

  }


  return text;

}
