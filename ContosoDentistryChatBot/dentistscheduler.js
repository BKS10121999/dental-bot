class DentistScheduler {
    constructor(configuration) {
        this.getAvailability = async () => {
            const response = await fetch(configuration.SchedulerEndpoint + "availability")
            const times = await response.json()
            let responseText = `Current time slots available: `
            times.map(time => {
                responseText += `
${time}`
            })
            return responseText
        }

        this.scheduleAppointment = async (time) => {
            const response = await fetch(configuration.SchedulerEndpoint + "schedule", {
                method: "post",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ time: time })
            })
            let responseText = `An appointment is set for ${time}.`
            return responseText
        }
    }
}

module.exports = DentistScheduler
