export const groupBy = (collection, column) => {
    return collection.reduce((x,y) => {
        (x[column(y)] = x[column(y)] || []).push(y);
        return x;
    }, {});
}