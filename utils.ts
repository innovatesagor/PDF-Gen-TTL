export const getFormattedDate = (date: Date): string => {
  // Returns format: 01-Jan-2025
  return new Intl.DateTimeFormat('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  }).format(date).replace(/ /g, '-');
};

export const getFilenameDate = (dateStr: string): string => {
  // Input: 01-Jan-2025 -> Output: 01.01.25
  try {
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return '00.00.00';
    
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = String(date.getFullYear()).slice(-2);
    
    return `${day}.${month}.${year}`;
  } catch (e) {
    return '00.00.00';
  }
};

export const formatCurrency = (amount: number): string => {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    minimumFractionDigits: 2
  }).format(amount);
};
