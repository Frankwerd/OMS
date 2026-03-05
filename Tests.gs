/********************************
 * Tests.gs
 * Manual verification suite for SKU derivation and Address Parsing
 ********************************/

function runAllTests() {
  test_deriveSku();
  test_parseGlobalAddress();
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
      input: ['Invalid address block no geo info'],
      expected: { addr1: 'Invalid address block no geo info', success: false }
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
