export const rgbToArgb = (rgb: string): string => {
    const numbers = rgb.split('(')[1].split(')')[0].split(',')
        .map((number, index) => {
            return index === 3 ? `${parseInt(number) * 255}` : number
        })
    if (numbers.length === 3) {
        numbers.unshift('255')
    }
    const argb = numbers.map(number => {
        const hex = parseInt(number).toString(16)
        return hex.length === 1 ? `0${hex}` : hex
    })
    return argb.join('').toUpperCase()
}
