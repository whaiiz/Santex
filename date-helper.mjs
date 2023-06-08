export const parseDate = (date) => {
    try {
        const dateParts = date.split('-');
        const dateObject = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
        return dateObject;
    } catch(Exception) {
        throw new Error('Error parsing date: ', date);
    }
}

export const monthNumberToString = (monthNumber) => {
    const months = {
        '0': 'Janeiro',
        '1': 'Fevereiro',
        '2': 'Março',
        '3': 'Abril',
        '4': 'Maio',
        '5': 'Junho',
        '6': 'Julho',
        '7': 'Agosto',
        '8': 'Setembro',
        '9': 'Outubro',
        '10': 'Novembro',
        '11': 'Dezembro',
    }

    return months[monthNumber];
}