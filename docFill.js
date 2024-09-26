
const docx4js = require("docx4js")
const ModelHandler = require("docx4js/lib/openxml/docx/model-handler").default
          
Object.defineProperties(global , {
    Blob : require("cross-blob"),
})

function collisionArea1D(A1, B1, A2, B2) {
    const startMin = Math.min(A1, A2)
    const endMax = Math.max(B1, B2)
    const minMaxDist = endMax - startMin
    const length1 = B1 - A1
    const length2 = B2 - A2
    const lengthSum = length1 + length2
    const collisionArea = lengthSum - minMaxDist
    return collisionArea
}

function untrim(str) {
    const NON_BREAKING_SPACE = '\xa0'
    return str.replace(" ", NON_BREAKING_SPACE)
}

function resolveSubfix(origin, pos, mutations) {
    if (mutations.length == 0) {
        return origin.substring(pos);
    }
    const [firstMutation, ...remainingMutations] = mutations
    let segment = ""
    let skip = 0
    let nextMutation = remainingMutations
    if (firstMutation.index === pos) {
        segment = firstMutation.with
        skip = firstMutation.length
    } else {
        let segmentLength = firstMutation.index - pos
        skip = segmentLength
        segment = origin.substring(pos, pos + segmentLength)
        nextMutation = mutations
    }
    return untrim(segment) + untrim(resolveSubfix(origin, pos + skip, nextMutation))
}

function buildTextFromMutations(origin, mutations) {
    return resolveSubfix(origin, 0, mutations)
}


class MyModelhandler extends ModelHandler {
    constructor(content, repMap) {
        super()
        this.repMap = repMap
        this.cursor = 0
        this.debug = {}
        this.runsReplacements = []
        this.paragReplacements = []
        this.content = content
        this.variables = []
        this.paragReplacements = this.buildReplacements(this.content)
        this.debug.content = content
        this.paragLength = content.length
    }

    getRunReplacements(posInRarag, length) {
        let runReplacements = []
        const runStart = posInRarag
        const runEnd = posInRarag + length
        for (let rep of this.paragReplacements) {
            const replacementStart = rep.index
            const replacementEnd = rep.index + rep.length
            const collisionArea = Math.max(collisionArea1D(runStart, runEnd, replacementStart, replacementEnd), 0)
            const collision = !!collisionArea
            if (collision) {
                runReplacements = [...runReplacements, rep]
            }
        }
        return runReplacements
    }
    getVariables()
    {
        return this.variables
    }
    buildReplacements(text) {
        let replacements = []
        var regExp = /\$\(([^)]+)\)/g
        var match;
        while ((match = regExp.exec(text)) != null) {
            const index = match.index
            const length = match[0].length
            const varName = match[1]
            this.variables.push(varName)
            let with_ = null
            if (this.repMap){
                with_ = this.repMap[varName]
            }
            let replacement = null
            if (with_) {
                replacement = {
                    index, length, with: with_
                }
            }
            if (replacement) {
                replacements = [...replacements, replacement]
            }
        }
        return replacements
    }

    moveReplacementToRunSpace(runStart, runEnd, rep) {
        const replacementStart = rep.index
        const replacementEnd = replacementStart + rep.length
        const runLength = runEnd - runStart + 1
        const isLastRunForReplacement = replacementEnd >= runStart && replacementEnd <= runEnd
        return {
            index: Math.max(0, rep.index - runStart),
            length: Math.min(collisionArea1D(runStart, runEnd, replacementStart, replacementEnd), runLength),
            with: isLastRunForReplacement ? rep.with : ""
        }
    }
}


export function docGetVars(inputPath)
{
    return docx4js.docx.load(inputPath).then(docx => {
        return new Promise((resolve, reject) => {
            const content = docx.officeDocument.content("w\\:t").text()
            let handler = new MyModelhandler(content, undefined)
            resolve(handler.getVariables())
        })
    })
}

export function docFill(inputPath, outputPath, map) {
    return docx4js.docx.load(inputPath).then(docx => {
        return new Promise((resolve, reject) => {
            const content = docx.officeDocument.content("w\\:t").text()
            let handler = new MyModelhandler(content, map)
            handler.on("r", function (a0, a1, a2) {
                const content = a2.doc.officeDocument.content(a1).text()
                const runReplacements = handler.getRunReplacements(this.cursor, content.length)
                const runStart = this.cursor
                const runLength = content.length
                const runEnd = runStart + runLength
                let runReplacementsInRunSpace = []
                for (let rep of runReplacements) {
                    let repInRunSpace = handler.moveReplacementToRunSpace(runStart, runEnd, rep)
                    runReplacementsInRunSpace = [...runReplacementsInRunSpace, repInRunSpace]
                }
                this.runsReplacements = [...this.runsReplacements, runReplacementsInRunSpace]
                this.cursor += content.length
            })
            docx.parse(handler)
            let runs = docx.officeDocument.content("w\\:t")
            for (let i = 0; i < runs.length; i++) {
                const text = docx.officeDocument.content(runs[i]).text()
                const final = buildTextFromMutations(text, handler.runsReplacements[i])
                docx.officeDocument.content(runs[i]).text(final)
            }
            docx.save(outputPath)
            resolve(undefined)
        })
    })
}

 