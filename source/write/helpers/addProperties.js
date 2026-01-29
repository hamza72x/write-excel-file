export default function addProperties(otherProperties, object, propertyNames) {
	for (const propertyName of propertyNames) {
		if (object[propertyName] !== undefined && object[propertyName] !== null) {
			otherProperties[propertyName] = object[propertyName]
		}
	}
}