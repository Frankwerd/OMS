/********************************
 * Tests.gs
 * Manual verification suite for SKU derivation and Address Parsing
 ********************************/

function runAllTests() {
  test_deriveSku();
  test_parseGlobalAddress();
  test_normalizers();
  test_migrationHelpers();
}

function test_deriveSku() {
  const cases = [
    {
      input: { model: 'Pro', clubType: 'Wood', hand: 'Left', flex: 'Regular', length: 'Standard', gripSize: 'Standard', magSafeStand: '0' },
      expected: 'GG-PRO-WD-LR-ST-ST-0'
    },
    {
      input: { model: 'Basic', clubType: 'Iron', hand: 'Right', flex: 'Stiff', length: 'Longer', gripSize: 'Mid', magSafeStand: 'Yes' },
      expected: 'GG-BAS-IR-RS-LG-MS-M'
    },
    {
      input: { model: 'Pro', clubType: 'Iron', hand: 'Left', flex: 'L-Flex', length: 'Standard', gripSize: 'Standard', magSafeStand: 'TRUE' },
      expected: 'GG-PRO-IR-LL-ST-ST-M'
    },
    {
      input: { model: 'Basic', clubType: 'Wood', hand: 'Right', flex: 'X-Stiff', length: 'Longer', gripSize: 'Standard', magSafeStand: '1' },
      expected: 'GG-BAS-WD-RX-LG-ST-M'
    },
    {
      input: {}, // defaults
      expected: 'GG-BAS-IR-RR-ST-ST-0'
    }
  ];

  cases.forEach((c, i) => {
    const result = OMS_Utils.deriveSku(c.input);
    if (result !== c.expected) {
      console.error(`test_deriveSku Case ${i} FAILED. Expected: ${c.expected}, Got: ${result}`);
    } else {
      console.log(`test_deriveSku Case ${i} PASSED.`);
    }
  });
}

function test_migrationHelpers() {
  // Currency
  if (migrate_parseCurrency_('$220.00') !== 220) console.error('test_migrationHelpers Currency 1 FAILED');
  if (migrate_parseCurrency_('2,220.50') !== 2220.50) console.error('test_migrationHelpers Currency 2 FAILED');
  
  // Phone
  if (migrate_parsePhone_('="3104302093"') !== '+1-310-430-2093') console.error('test_migrationHelpers Phone 1 FAILED');
  
  // Date
  const d1 = migrate_parseDate_('02/17/2026');
  if (!d1 || d1.getFullYear() !== 2026 || d1.getMonth() !== 1 || d1.getDate() !== 17) {
    console.error('test_migrationHelpers Date 1 FAILED', d1);
  }
  
  // Completion Date
  const cd1 = migrate_parseCompletionDate_('Wednesday, 2/25/26 at 11:55 AM');
  if (!cd1 || cd1.getFullYear() !== 2026 || cd1.getMonth() !== 1 || cd1.getDate() !== 25) {
    console.error('test_migrationHelpers Completion Date 1 FAILED', cd1);
  }
  
  console.log('test_migrationHelpers PASSED (if no errors above).');
}

function test_normalizers() {
  // Email
  const e1 = OMS_Utils.normalizeEmail_('  Test@Example.Com  ');
  if (e1 !== 'test@example.com') console.error(`test_normalizers Email FAILED. Got: ${e1}`);
  else console.log('test_normalizers Email PASSED.');

  // Phone
  const cases = [
    { input: '4848948972', expected: '+1-484-894-8972' },
    { input: '14848948972', expected: '+1-484-894-8972' },
    { input: '+1 (123) 456-7890', expected: '+1-123-456-7890' },
    { input: '9999', expected: '9999' }
  ];

  cases.forEach((c, i) => {
    const result = OMS_Utils.normalizePhone(c.input);
    if (result !== c.expected) {
      console.error(`test_normalizers Phone Case ${i} FAILED. Expected: ${c.expected}, Got: ${result}`);
    } else {
      console.log(`test_normalizers Phone Case ${i} PASSED.`);
    }
  });
}

function test_parseGlobalAddress() {
  const cases = [
    {
      input: ['123 Main St', 'San Francisco CA 94105', 'USA'],
      expected: { addr1: '123 Main St', city: 'San Francisco', state: 'CA', zip: '94105', country: 'United States', success: true }
    },
    {
      input: ['10 Downing St', 'London SW1A 2AA', 'United Kingdom'],
      expected: { addr1: '10 Downing St', city: 'London', zip: 'SW1A 2AA', country: 'United Kingdom', success: true }
    },
    {
      input: ['[12345] 123 Gangnam-daero', 'Seoul', 'South Korea'],
      expected: { addr1: '123 Gangnam-daero', city: 'Seoul', zip: '12345', country: 'South Korea', success: true }
    },
    {
      input: ['1-1 Chiyoda', 'Chiyoda-ku, Tokyo 100-8111', 'Japan'],
      expected: { addr1: '1-1 Chiyoda', city: 'Chiyoda-ku, Tokyo', zip: '100-8111', country: 'Japan', success: true }
    },
    {
      input: ['Rue de la Paix 1', '75002 Paris', 'France'],
      expected: { addr1: 'Rue de la Paix 1', city: 'Paris', zip: '75002', country: 'France', success: true }
    },
    {
      input: ['123 Austin Way', 'Austin Texas 78701', 'USA'],
      expected: { addr1: '123 Austin Way', city: 'Austin', state: 'TX', zip: '78701', country: 'United States', success: true }
    },
    {
      input: ['456 Reversed Rd', '90210 Beverly Hills CA', 'USA'],
      expected: { addr1: '456 Reversed Rd', city: 'Beverly Hills', state: 'CA', zip: '90210', country: 'United States', success: true }
    },
    {
      input: ['Center valley, PA 18034 United States'],
      expected: { addr1: '', city: 'Center valley', state: 'PA', zip: '18034', country: 'United States', success: true }
    },
    {
      input: ['Invalid address block no geo info'],
      expected: { addr1: 'Invalid address block no geo info', city: '', state: '', zip: '', success: false }
    }
  ];

  cases.forEach((c, i) => {
    const result = OMS_Utils.parseGlobalAddress(c.input);
    const pass = (result.addr1 === c.expected.addr1 &&
                  result.country === (c.expected.country || 'United States') &&
                  result.success === c.expected.success &&
                  (c.expected.zip ? result.zip === c.expected.zip : true) &&
                  (c.expected.city ? result.city === c.expected.city : true));

    if (!pass) {
      console.error(`test_parseGlobalAddress Case ${i} FAILED. Got: ${JSON.stringify(result)}`);
    } else {
      console.log(`test_parseGlobalAddress Case ${i} PASSED.`);
    }
  });
}
