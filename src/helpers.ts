export const has_duplicates = (arr: any[], isSilent: boolean): boolean => {
  const seen = new Set();
  for (const item of arr) {
    if (item !== '' && item !== null && item !== undefined) {
      if (seen.has(item)) {
        to_console(`❗️ Duplicate value: ${item}`, isSilent);
        return true;
      }
      seen.add(item);
    }
  }
  return false;
};

export const to_console = (msg: string, isSilent: boolean) => {
  if (!isSilent) {
    const prefix = '[Google Sheets Lib]';
    console.log(`${prefix} ${msg}`);
  }
};
